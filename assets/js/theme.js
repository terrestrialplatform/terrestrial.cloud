// ==========================================================================
// Terrestrial Platform - Global Theme Manager (Dark / Light Theme)
// Persists choice in localStorage across all pages with zero FOUC
// ==========================================================================

(function() {
    const THEME_KEY = 'terra_theme_pref';
    const urlParams = new URLSearchParams(window.location.search);
    const urlTheme = urlParams.get('theme');
    const saved = (urlTheme === 'light' || urlTheme === 'dark') ? urlTheme : localStorage.getItem(THEME_KEY);
    
    // Default to light for main homepage if not set; if set, respect saved preference
    let initialTheme = (saved === 'light' || saved === 'dark') ? saved : 'light';
    
    document.documentElement.setAttribute('data-theme', initialTheme);

    window.toggleTheme = function() {
        const current = document.documentElement.getAttribute('data-theme') || 'light';
        const next = current === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', next);
        try {
            localStorage.setItem(THEME_KEY, next);
        } catch (e) {
            console.warn('localStorage access denied for theme preference');
        }
        updateThemeToggleUI(next);
    };

    function updateThemeToggleUI(theme) {
        const buttons = document.querySelectorAll('.theme-toggle-btn');
        buttons.forEach(btn => {
            const label = btn.querySelector('.theme-label');
            if (label) {
                label.textContent = theme === 'dark' ? 'Dark' : 'Light';
            }
            btn.setAttribute('aria-label', `Current theme: ${theme}. Click to switch to ${theme === 'dark' ? 'light' : 'dark'} mode`);
            btn.setAttribute('title', `Switch to ${theme === 'dark' ? 'Light' : 'Dark'} Theme`);
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() {
            const current = document.documentElement.getAttribute('data-theme') || 'light';
            updateThemeToggleUI(current);
        });
    } else {
        const current = document.documentElement.getAttribute('data-theme') || 'light';
        updateThemeToggleUI(current);
    }
})();
