// ==========================================================================
// Terrestrial Platform - Global Theme Manager (Enforced Light Theme)
// ==========================================================================

(function() {
    // 1. Immediate theme initialization to light mode
    document.documentElement.setAttribute('data-theme', 'light');
    try {
        localStorage.setItem('terra_theme_pref', 'light');
    } catch (e) {
        // Storage unavailable
    }

    // 2. Dummy toggle to maintain compatibility if called
    window.toggleTheme = function() {
        document.documentElement.setAttribute('data-theme', 'light');
    };

    // 3. Dropdown toggle function
    window.toggleAppMenu = function(event) {
        if (event) {
            event.stopPropagation();
        }
        const m = document.getElementById('appMenu');
        const b = document.querySelector('.dropdown-toggle');
        if (m) {
            m.classList.toggle('open');
        }
        if (b) {
            b.classList.toggle('active');
        }
    };

    // 4. Close dropdown on outside click
    document.addEventListener('click', function(e) {
        const m = document.getElementById('appMenu');
        const b = document.querySelector('.dropdown-toggle');
        if (m && !e.target.closest('.dropdown-container')) {
            m.classList.remove('open');
            if (b) {
                b.classList.remove('active');
            }
        }
    });

    // 5. Global login opener
    window.btnNT = function() {
        window.open('https://platform.innovativeengineers.co.in/platform/', '_blank');
    };
})();
