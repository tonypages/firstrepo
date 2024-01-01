// Check if the user agent indicates a Windows desktop
function isWindowsDesktop() {
    const userAgent = window.navigator.userAgent;
    return /*/Windows NT/.test(userAgent) && */!/Mobile/.test(userAgent);
}

// Redirect if not a Windows desktop
function redirectToSSSPage() {
    const redirectUrl = document.body.getAttribute('data-redirect-url');
    if (!isWindowsDesktop()) {
        window.location.href = redirectUrl;
    }
}

// Call the function before the page loads
redirectToSSSPage();