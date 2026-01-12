const TOKEN_KEY = 'pxm_auth_token';
const USER_KEY = 'pxm_auth_user';
const ROLE_KEY = 'pxm_auth_role';
const API_BASE_URL = `http://${window.location.hostname}:3001/api`;

export function getAccessToken() {
  return localStorage.getItem(TOKEN_KEY);
}

export function getUser() {
  return localStorage.getItem(USER_KEY);
}

export function getRole() {
  return localStorage.getItem(ROLE_KEY);
}

export async function login(username, password) {
  const response = await fetch(`${API_BASE_URL}/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password })
  });

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.message || 'Anmeldung fehlgeschlagen');
  }

  const data = await response.json();
  localStorage.setItem(TOKEN_KEY, data.token);
  localStorage.setItem(USER_KEY, data.user);
  localStorage.setItem(ROLE_KEY, data.role);
  return { user: data.user, role: data.role };
}

export function logout() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
  localStorage.removeItem(ROLE_KEY);
}