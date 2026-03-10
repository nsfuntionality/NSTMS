// Check if already logged in
(function checkAuth() {
    const user = sessionStorage.getItem('tms_user');
    if (user) {
        window.location.href = 'tms.html';
    }
})();

document.getElementById('loginForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value.trim();
    const errorEl = document.getElementById('loginError');
    errorEl.style.display = 'none';

    fetch('data/settings.json')
        .then(function(res) { return res.json(); })
        .then(function(settings) {
            const user = settings.users.find(function(u) {
                return u.username === username && u.password === password;
            });
            if (user) {
                sessionStorage.setItem('tms_user', JSON.stringify({ username: user.username, name: user.name }));
                // Store audit mode setting
                sessionStorage.setItem('tms_audit_mode', settings.auditMode ? 'true' : 'false');
                window.location.href = 'tms.html';
            } else {
                errorEl.textContent = 'Invalid username or password.';
                errorEl.style.display = 'block';
            }
        })
        .catch(function() {
            errorEl.textContent = 'Unable to connect. Please try again.';
            errorEl.style.display = 'block';
        });
});
