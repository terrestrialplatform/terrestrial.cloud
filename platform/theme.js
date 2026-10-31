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

    // 6. Mobile navigation toggle function
    window.toggleMobileNav = function(event) {
        if (event) {
            event.stopPropagation();
        }
        const nav = document.querySelector('.corp-nav-links');
        const btn = document.querySelector('.mobile-menu-btn');
        if (nav) {
            nav.classList.toggle('mobile-open');
        }
        if (btn) {
            btn.classList.toggle('active');
        }
    };

    // 7. Auto-inject mobile hamburger button if not present in .nav-wrap
    function initMobileNav() {
        const navWrap = document.querySelector('.nav-wrap');
        if (navWrap && !document.querySelector('.mobile-menu-btn')) {
            const btn = document.createElement('button');
            btn.className = 'mobile-menu-btn';
            btn.setAttribute('type', 'button');
            btn.setAttribute('aria-label', 'Toggle Navigation Menu');
            btn.innerHTML = '<span></span><span></span><span></span>';
            btn.onclick = window.toggleMobileNav;
            navWrap.appendChild(btn);
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initMobileNav);
    } else {
        initMobileNav();
    }

    // 8. Close mobile nav on outside click
    document.addEventListener('click', function(e) {
        const nav = document.querySelector('.corp-nav-links');
        const btn = document.querySelector('.mobile-menu-btn');
        if (nav && nav.classList.contains('mobile-open')) {
            if (!e.target.closest('.corp-nav-links') && !e.target.closest('.mobile-menu-btn')) {
                nav.classList.remove('mobile-open');
                if (btn) {
                    btn.classList.remove('active');
                }
            }
        }
    });
})();
