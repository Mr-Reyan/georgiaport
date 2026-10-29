// Port Pages JavaScript - Centralized for all 49 port pages

/* ---------- Ports data ---------- */
const ports = [
  "Ship Chandler Dar Es Salaam", "Ship Chandler Apapa", "Ship Chandler Onne Port Nigeria", "San Pedro Ship Supply Chandler",
  "Ship Chandler Abidjan", "Ship Chandler Freetown", "Monrovia Ship Chandler", "Ship Chandler Maputo", "Ship Chandler Matadi",
  "Ship Chandler Beira", "Ship Chandler Angola", "Ship Chandler Lobito", "ShipChandler Lome", "Shipchandler Lome",
  "Ivory Coast Ship Chandler", "Nouadhibou Ship Chandler", "Ship Chandler Luanda", "Ship Chandler Cape Verde",
  "Ship Chandler Gabon", "Ship Chandler Namibia", "Walvis Bay Ship Chandler", "Ship Chandler Gentil Port - Gabon",
  "Ship Chandler Takoradi", "Ship Chandler Congo", "Shipchandler Lagos", "Ship Chandler Port Ambon",
  "Shipchandler Pointe Noire", "Ship Chandler Djibouti", "Ship Chandler Cabinda", "Ship Chandler Tema",
  "Ship Chandler Mindelo", "Ship Chandler Casablanca", "Ship Chandler Safi Port Morocco", "Ship Chandler Kamsar",
  "Ship Chandler Soyo", "Dakar Ship Chandler", "Ship Chandler Senegal", "Ship Chandler Dakar", "Bissau Ship Chandler - Guinea",
  "Conakry Ship Chandler", "Ship Chandler Cotonou", "Douala Ship Chandler", "Liberia Ship Chandler", "Namibia Ship Chandler",
  "Ship Chandler Dakar", "Ship Chandler Mombasa", "Ship Chandlers Saldanha Bay", "Ship Chandler Owendo", "Ship Chandlers Richards Bay"
];

const megaPorts = document.getElementById('megaPorts');
if (megaPorts) {
  const third = Math.ceil(ports.length / 3);
  const cols = [ports.slice(0, third), ports.slice(third, third * 2), ports.slice(third * 2)];
  const titles = ['African Ports (A-D)', 'African Ports (E-M)', 'African Ports (N-Z)'];

  cols.forEach((list, i) => {
    const col = document.createElement('div');
    col.className = 'mega-col';
    col.innerHTML = `<h5>${titles[i]}</h5>`;
    list.forEach(port => {
      const a = document.createElement('a');
      a.href = `${port.toLowerCase().replace(/[^a-z0-9]+/g, '-')}.html`;
      a.textContent = port;
      col.appendChild(a);
    });
    megaPorts.appendChild(col);
  });

  const portSearch = document.getElementById('portSearch');
  if (portSearch) {
    portSearch.addEventListener('input', function (e) {
      const q = e.target.value.toLowerCase();
      const links = megaPorts.querySelectorAll('a');
      const groups = megaPorts.querySelectorAll('.mega-col');

      links.forEach(l => {
        l.style.display = l.textContent.toLowerCase().includes(q) ? 'block' : 'none';
      });

      groups.forEach(g => {
        const any = Array.from(g.querySelectorAll('a')).some(a => a.style.display !== 'none');
        g.style.display = any ? 'block' : 'none';
      });
    });
  }
}

document.addEventListener('DOMContentLoaded', function () {

  /* Sticky header */
  const masthead = document.getElementById('masthead');
  if (masthead) {
    window.addEventListener('scroll', () => {
      masthead.classList.toggle('is-stuck', window.scrollY > 20);
    });
  }

  /* Burger */
  const burger = document.querySelector('.burger');
  const navShell = document.getElementById('navShell');

  if (burger && navShell && masthead) {
    burger.addEventListener('click', function () {
      navShell.classList.toggle('open');
      this.classList.toggle('open');
      masthead.classList.toggle('is-open');
    });
  }

  /* Mobile dropdowns */
  document.querySelectorAll('.nav-unit.has-flyout, .nav-unit.has-mega').forEach(unit => {
    const anchor = unit.querySelector('.nav-anchor');
    if (anchor) {
      anchor.addEventListener('click', function (e) {
        if (window.innerWidth <= 1024) {
          e.preventDefault();
          unit.classList.toggle('open');
        }
      });
    }
  });

  /* JS Fallback for dropdown flicker (guarantees smooth hover) */
  const navPorts = document.getElementById('navPorts');
  const megaMenu = document.getElementById('megaPortsMenu');

  if (navPorts && megaMenu) {
    let timeout;
    const openMenu = () => {
      clearTimeout(timeout);
      megaMenu.style.opacity = '1';
      megaMenu.style.visibility = 'visible';
      megaMenu.style.transform = 'translateX(-50%) translateY(0)';
    };
    const closeMenu = () => {
      timeout = setTimeout(() => {
        megaMenu.style.opacity = '0';
        megaMenu.style.visibility = 'hidden';
        megaMenu.style.transform = 'translateX(-50%) translateY(14px)';
      }, 100);
    };

    if (window.innerWidth > 1024) {
      navPorts.parentElement.addEventListener('mouseenter', openMenu);
      navPorts.parentElement.addEventListener('mouseleave', closeMenu);
      megaMenu.addEventListener('mouseenter', openMenu);
      megaMenu.addEventListener('mouseleave', closeMenu);
    }
  }

  /* Scroll reveal */
  const risers = document.querySelectorAll('.rise');
  if (risers.length > 0) {
    const io = new IntersectionObserver((entries) => {
      entries.forEach(en => {
        if (en.isIntersecting) {
          en.target.classList.add('in');
          io.unobserve(en.target);
        }
      });
    }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });

    risers.forEach(el => io.observe(el));
  }

  /* Horizontal TOC Active State */
  const tocLinks = document.querySelectorAll('.port-toc-bar a');
  const categories = document.querySelectorAll('.service-category');

  if (tocLinks.length > 0 && categories.length > 0) {
    window.addEventListener('scroll', () => {
      let current = '';
      const scrollPos = window.scrollY + 160;

      categories.forEach(cat => {
        const top = cat.offsetTop;
        if (scrollPos >= top) {
          current = cat.getAttribute('id');
        }
      });

      tocLinks.forEach(link => {
        link.classList.remove('active');
        if (current && link.getAttribute('href') === `#${current}`) {
          link.classList.add('active');
        }
      });
    });

    /* Smooth scroll for TOC */
    tocLinks.forEach(link => {
      link.addEventListener('click', function (e) {
        e.preventDefault();
        const targetId = this.getAttribute('href');
        const targetEl = document.querySelector(targetId);
        if (targetEl) {
          const headerOffset = 140;
          const elementPosition = targetEl.getBoundingClientRect().top;
          const offsetPosition = elementPosition + window.pageYOffset - headerOffset;
          window.scrollTo({
            top: offsetPosition,
            behavior: 'smooth'
          });
        }
      });
    });
  }
});
