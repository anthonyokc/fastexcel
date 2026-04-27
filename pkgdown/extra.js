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

  const list = document.createElement("ul");
  list.className = "list-unstyled";

  headings.forEach(function (heading) {
    const item = document.createElement("li");
    const link = document.createElement("a");

    link.href = "#" + heading.id;
    link.textContent = heading.textContent.trim();

    item.appendChild(link);
    list.appendChild(item);
  });

  nav.appendChild(list);
  aside.appendChild(nav);
});
