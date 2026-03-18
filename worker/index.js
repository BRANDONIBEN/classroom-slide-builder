// Cloudflare Worker — GitHub API proxy for Classroom Slide Builder
// Keeps GITHUB_TOKEN server-side. Webapp calls this instead of GitHub directly.

const GITHUB_API = 'https://api.github.com';

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';

    // CORS preflight
    if (request.method === 'OPTIONS') {
      return corsResponse(env, new Response(null, { status: 204 }));
    }

    try {
      let response;
      const path = url.pathname;

      if (path === '/health') {
        response = json({ status: 'ok' });
      } else if (path === '/courses' && request.method === 'GET') {
        response = await getCourseIndex(env);
      } else if (path.match(/^\/courses\/([a-z0-9_-]+)$/) && request.method === 'GET') {
        const id = path.split('/')[2];
        response = await getCourse(env, id);
      } else if (path.match(/^\/courses\/([a-z0-9_-]+)$/) && request.method === 'PUT') {
        const id = path.split('/')[2];
        const body = await request.json();
        response = await saveCourse(env, id, body);
      } else if (path.match(/^\/courses\/([a-z0-9_-]+)\/history$/) && request.method === 'GET') {
        const id = path.split('/')[2];
        response = await getCourseHistory(env, id);
      } else if (path.match(/^\/courses\/([a-z0-9_-]+)\/revert\/(.+)$/) && request.method === 'POST') {
        const parts = path.split('/');
        const id = parts[2];
        const sha = parts[4];
        response = await revertCourse(env, id, sha);
      } else {
        response = json({ error: 'Not found' }, 404);
      }

      return corsResponse(env, response);
    } catch (err) {
      return corsResponse(env, json({ error: err.message }, 500));
    }
  }
};

// --- GitHub API helpers ---

async function ghFetch(env, path, options = {}) {
  const url = `${GITHUB_API}/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${path}?ref=${env.GITHUB_BRANCH}`;
  const resp = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${env.GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'classroom-api-worker',
      ...(options.headers || {})
    }
  });
  return resp;
}

// GET /courses — returns data/index.json
async function getCourseIndex(env) {
  const resp = await ghFetch(env, 'data/index.json');
  if (!resp.ok) return json({ error: 'Failed to load index', status: resp.status }, 502);
  const data = await resp.json();
  // GitHub returns base64 content
  const content = JSON.parse(atob(data.content));
  return json(content);
}

// GET /courses/:id — returns data/{id}_text.json
async function getCourse(env, id) {
  const resp = await ghFetch(env, `data/${id}_text.json`);
  if (!resp.ok) return json({ error: 'Course not found: ' + id }, 404);
  const data = await resp.json();
  const content = JSON.parse(atob(data.content));
  return json(content);
}

// PUT /courses/:id — commits updated JSON to data/{id}_text.json
async function saveCourse(env, id, body) {
  const filePath = `data/${id}_text.json`;

  // Get current file SHA (required for update)
  const existing = await ghFetch(env, filePath);
  let sha = null;
  if (existing.ok) {
    const data = await existing.json();
    sha = data.sha;
  }

  // Commit the file
  const content = btoa(unescape(encodeURIComponent(JSON.stringify(body, null, 2))));
  const commitBody = {
    message: `Update ${id} course data`,
    content: content,
    branch: env.GITHUB_BRANCH
  };
  if (sha) commitBody.sha = sha;

  const url = `${GITHUB_API}/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${filePath}`;
  const resp = await fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${env.GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'classroom-api-worker',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(commitBody)
  });

  if (!resp.ok) {
    const err = await resp.text();
    return json({ error: 'GitHub commit failed', detail: err }, 502);
  }

  const result = await resp.json();
  return json({ ok: true, sha: result.content.sha, commit: result.commit.sha });
}

// GET /courses/:id/history — last 20 commits for the course file
async function getCourseHistory(env, id) {
  const filePath = `data/${id}_text.json`;
  const url = `${GITHUB_API}/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/commits?path=${filePath}&per_page=20&sha=${env.GITHUB_BRANCH}`;
  const resp = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${env.GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'classroom-api-worker'
    }
  });
  if (!resp.ok) return json({ error: 'Failed to load history' }, 502);
  const commits = await resp.json();
  const history = commits.map(c => ({
    sha: c.sha,
    message: c.commit.message,
    date: c.commit.author.date,
    author: c.commit.author.name
  }));
  return json({ history });
}

// POST /courses/:id/revert/:sha — restore file to a previous commit's version
async function revertCourse(env, id, sha) {
  const filePath = `data/${id}_text.json`;
  // Get the file at the old commit
  const oldUrl = `${GITHUB_API}/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${filePath}?ref=${sha}`;
  const oldResp = await fetch(oldUrl, {
    headers: {
      'Authorization': `Bearer ${env.GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'classroom-api-worker'
    }
  });
  if (!oldResp.ok) return json({ error: 'Version not found' }, 404);
  const oldData = await oldResp.json();

  // Get current file SHA
  const curResp = await fetch(`${GITHUB_API}/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${filePath}?ref=${env.GITHUB_BRANCH}`, {
    headers: {
      'Authorization': `Bearer ${env.GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'classroom-api-worker'
    }
  });
  const curData = curResp.ok ? await curResp.json() : null;

  // Commit the old content as new
  const commitBody = {
    message: `Revert ${id} to version from ${sha.substring(0, 7)}`,
    content: oldData.content.replace(/\n/g, ''),
    branch: env.GITHUB_BRANCH
  };
  if (curData && curData.sha) commitBody.sha = curData.sha;

  const putResp = await fetch(`${GITHUB_API}/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${filePath}`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${env.GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'classroom-api-worker',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(commitBody)
  });

  if (!putResp.ok) {
    const err = await putResp.text();
    return json({ error: 'Revert failed', detail: err }, 502);
  }
  return json({ ok: true, message: 'Reverted successfully' });
}

// --- Utility ---

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}

function corsResponse(env, response) {
  const headers = new Headers(response.headers);
  headers.set('Access-Control-Allow-Origin', env.ALLOWED_ORIGIN || '*');
  headers.set('Access-Control-Allow-Methods', 'GET, PUT, OPTIONS');
  headers.set('Access-Control-Allow-Headers', 'Content-Type');
  headers.set('Access-Control-Max-Age', '86400');
  return new Response(response.body, {
    status: response.status,
    headers
  });
}
