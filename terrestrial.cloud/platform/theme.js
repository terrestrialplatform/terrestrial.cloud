// ==========================================================================
// Terrestrial Platform - Global Theme Manager (Dark / Light Theme)
// Persists choice in localStorage across all pages with zero FOUC
// ==========================================================================

(function() {
    // 1. Immediate theme initialization before DOM content renders to prevent flash
    const THEME_KEY = 'terra_theme_pref';
    const saved = localStorage.getItem(THEME_KEY);
    
    // Default to dark theme if no preference is saved
    let initialTheme = (saved === 'light' || saved === 'dark') ? saved : 'dark';
    
    document.documentElement.setAttribute('data-theme', initialTheme);

    // 2. Global toggle function exposed to window
    window.toggleTheme = function() {
        const current = document.documentElement.getAttribute('data-theme') || 'dark';
        const next = current === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', next);
        try {
            localStorage.setItem(THEME_KEY, next);
        } catch (e) {
            console.warn('localStorage access denied for theme preference');
        }
        updateThemeToggleUI(next);
    };

    // 3. UI sync function for button labels, icons, and tooltips
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

    // 4. Sync dropdown toggle button label with active product name
    function syncDropdownAppName() {
        const activeLink = document.querySelector('#appMenu .dropdown-link.active span:first-child');
        const toggleSpan = document.querySelector('.dropdown-toggle span');
        if (activeLink && toggleSpan) {
            toggleSpan.textContent = activeLink.textContent.trim();
        }
    }

    // 5. Update UI once DOM is loaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() {
            const current = document.documentElement.getAttribute('data-theme') || 'dark';
            updateThemeToggleUI(current);
            syncDropdownAppName();
        });
    } else {
        const current = document.documentElement.getAttribute('data-theme') || 'dark';
        updateThemeToggleUI(current);
        syncDropdownAppName();
    }
})();
