document.addEventListener("DOMContentLoaded", function () {
  const page = document.querySelector(".template-home");
  const aside = page && page.querySelector("aside");
  const main = page && page.querySelector("main");

  if (!aside || !main || aside.querySelector("#toc")) {
    return;
  }

  const headings = Array.from(main.querySelectorAll("h2[id]"));

  if (headings.length === 0) {
    return;
  }

  const nav = document.createElement("nav");
  nav.id = "toc";
  nav.setAttribute("aria-label", "Table of contents");

  const title = document.createElement("h2");
  title.textContent = "On this page";
  title.setAttribute("data-toc-skip", "");
  nav.appendChild(title);

  aside.appendChild(nav);

  const initHomeTocTracking = function (attempt) {
    const links = Array.from(nav.querySelectorAll('a.nav-link[href^="#"]'));

    if (links.length === 0) {
      if (attempt < 10) {
        window.requestAnimationFrame(function () {
          initHomeTocTracking(attempt + 1);
        });
      }
      return;
    }

    const setActiveLink = function () {
      const threshold = 96;
      let activeId = headings[0].id;

      headings.forEach(function (heading) {
        if (heading.getBoundingClientRect().top <= threshold) {
          activeId = heading.id;
        }
      });

      links.forEach(function (link) {
        const id = decodeURIComponent(link.getAttribute("href").slice(1));
        const isActive = id === activeId;

        link.classList.toggle("active", isActive);

        if (isActive) {
          link.setAttribute("aria-current", "true");
        } else {
          link.removeAttribute("aria-current");
        }
      });
    };

    setActiveLink();
    window.addEventListener("scroll", setActiveLink, { passive: true });
    window.addEventListener("resize", setActiveLink);
  };

  initHomeTocTracking(0);
});
