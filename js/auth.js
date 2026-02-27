// Dummy authentication
const DUMMY_USERS = [
    { username: 'admin', password: 'admin123', name: 'Admin' },
    { username: 'dispatch', password: 'dispatch123', name: 'Dispatcher' },
    { username: 'driver', password: 'driver123', name: 'Driver' }
];

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

    const user = DUMMY_USERS.find(u => u.username === username && u.password === password);
    if (user) {
        sessionStorage.setItem('tms_user', JSON.stringify({ username: user.username, name: user.name }));
        window.location.href = 'tms.html';
    } else {
        alert('Invalid username or password.\n\nTry: admin / admin123');
    }
});
