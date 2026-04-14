function ouvrirLogin() {
  document.getElementById('page-login').classList.add('actif');
  setTimeout(function() { document.getElementById('login-id').focus(); }, 100);
}
function fermerLogin() {
  document.getElementById('page-login').classList.remove('actif');
  document.getElementById('login-erreur').style.display = 'none';
  document.getElementById('login-id').value = '';
  document.getElementById('login-pwd').value = '';
}
async function verifierLogin() {
  var id  = document.getElementById('login-id').value.trim();
  var pwd = document.getElementById('login-pwd').value;
  var err = document.getElementById('login-erreur');
  try {
    const r = await fetch('/api/login', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({username: id, password: pwd})
    });
    const data = await r.json();
    if (data.success && data.token) {
      sessionStorage.setItem('gavroche_token', data.token);
      err.style.display = 'none';
      window.location.href = '/admin.html';
    } else {
      err.style.display = 'block';
    }
  } catch(e) {
    err.style.display = 'block';
  }
}
