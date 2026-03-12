/* ==========================================================================
   Meenal Bagla Campaign Site – Main JavaScript
   ========================================================================== */

(function () {
  'use strict';

  // ── NAVIGATION ────────────────────────────────────────────────────────────

  const nav       = document.getElementById('nav');
  const navToggle = document.getElementById('navToggle');
  const navLinks  = document.getElementById('navLinks');

  // Sticky nav style on scroll
  window.addEventListener('scroll', () => {
    nav.classList.toggle('scrolled', window.scrollY > 60);
  }, { passive: true });

  // Mobile menu toggle
  navToggle.addEventListener('click', () => {
    navLinks.classList.toggle('open');
    const isOpen = navLinks.classList.contains('open');
    navToggle.setAttribute('aria-expanded', isOpen);
    // Animate hamburger → X
    navToggle.classList.toggle('active', isOpen);
  });

  // Close menu on nav link click (mobile)
  navLinks.querySelectorAll('.nav__link').forEach(link => {
    link.addEventListener('click', () => {
      navLinks.classList.remove('open');
      navToggle.classList.remove('active');
    });
  });

  // Active nav link on scroll
  const sections = document.querySelectorAll('section[id]');
  const linkEls  = navLinks.querySelectorAll('.nav__link[href^="#"]');

  function updateActiveLink() {
    const scrollY = window.scrollY + 100;
    let current = '';
    sections.forEach(section => {
      if (scrollY >= section.offsetTop) current = section.id;
    });
    linkEls.forEach(link => {
      link.classList.toggle('active', link.getAttribute('href') === `#${current}`);
    });
  }
  window.addEventListener('scroll', updateActiveLink, { passive: true });

  // ── FORM HANDLING ─────────────────────────────────────────────────────────

  const form       = document.getElementById('contactForm');
  const submitBtn  = document.getElementById('submitBtn');
  const successEl  = document.getElementById('formSuccess');
  const errorEl    = document.getElementById('formError');

  if (form) {
    form.addEventListener('submit', async (e) => {
      e.preventDefault();

      // Basic client-side validation
      const requiredFields = form.querySelectorAll('[required]');
      let valid = true;
      requiredFields.forEach(field => {
        field.classList.remove('error');
        if (!field.value.trim()) {
          field.classList.add('error');
          valid = false;
        }
      });
      if (!valid) {
        const firstError = form.querySelector('.error');
        if (firstError) firstError.focus();
        return;
      }

      // Email format check
      const emailField = form.querySelector('#email');
      const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (emailField && !emailRe.test(emailField.value.trim())) {
        emailField.classList.add('error');
        emailField.focus();
        return;
      }

      // Show loading state
      const btnText    = submitBtn.querySelector('.btn__text');
      const btnSpinner = submitBtn.querySelector('.btn__spinner');
      submitBtn.disabled = true;
      btnText.textContent = 'Sending…';

      // Build mailto fallback URL (always works, no server needed)
      const data = new FormData(form);
      const firstName = data.get('firstName') || '';
      const lastName  = data.get('lastName')  || '';
      const email     = data.get('email')     || '';
      const interest  = data.get('interest')  || '';
      const street    = data.get('street')    || '';
      const message   = data.get('message')   || '';

      const subject = encodeURIComponent(`Campaign Interest: ${firstName} ${lastName}`);
      const body = encodeURIComponent(
        `Name: ${firstName} ${lastName}\n` +
        `Email: ${email}\n` +
        `Street: ${street || 'Not provided'}\n` +
        `Interest: ${interest}\n\n` +
        `Message:\n${message || 'None'}\n\n` +
        `—\nSubmitted from meenalbagla.com`
      );

      // Try Formspree first; fall back gracefully to mailto
      try {
        const response = await fetch(form.action, {
          method: 'POST',
          body: data,
          headers: { 'Accept': 'application/json' },
        });

        if (response.ok) {
          showSuccess();
        } else {
          // Fallback: open mailto
          openMailto(subject, body);
          showSuccess();
        }
      } catch (_) {
        // Network error: open mailto
        openMailto(subject, body);
        showSuccess();
      }
    });

    // Remove error class on input
    form.querySelectorAll('.form-input').forEach(input => {
      input.addEventListener('input', () => input.classList.remove('error'));
    });
  }

  function openMailto(subject, body) {
    window.location.href = `mailto:meenalbagla@gmail.com?subject=${subject}&body=${body}`;
  }

  function showSuccess() {
    form.querySelectorAll('.form-group, .btn--submit, .form-note').forEach(el => {
      el.style.display = 'none';
    });
    const titleEl = form.querySelector('.contact__form-title');
    if (titleEl) titleEl.style.display = 'none';
    successEl.hidden = false;
    errorEl.hidden   = true;
    // Scroll success into view
    successEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  // ── INTERSECTION OBSERVER (subtle entrance animations) ────────────────────

  if ('IntersectionObserver' in window) {
    const observerOptions = {
      threshold: 0.12,
      rootMargin: '0px 0px -60px 0px',
    };

    const observer = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('in-view');
          observer.unobserve(entry.target);
        }
      });
    }, observerOptions);

    // Observe cards and key blocks
    document.querySelectorAll('.credential, .priority-card, .contact__way, .belmont-stat').forEach(el => {
      el.classList.add('observe');
      observer.observe(el);
    });
  }

  // ── SMOOTH SCROLL POLYFILL for older browsers ─────────────────────────────

  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
      const target = document.querySelector(this.getAttribute('href'));
      if (!target) return;
      e.preventDefault();
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });

})();
