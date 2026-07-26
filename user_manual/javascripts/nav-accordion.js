(() => {
  let initialized = false;

  const initAccordion = () => {
  const sidebar = document.querySelector(".md-sidebar--primary");
  if (!sidebar) {
    return;
  }

  const topLevelSections = Array.from(
    sidebar.querySelectorAll(
      ".md-nav--primary > .md-nav__list > .md-nav__item--section.md-nav__item--nested"
    )
  );

  if (!topLevelSections.length) {
    return;
  }

  const setExpanded = (section, expanded) => {
    const toggle = section.querySelector(":scope > input.md-nav__toggle");
    const nav = section.querySelector(":scope > nav.md-nav");
    if (toggle) {
      toggle.checked = expanded;
    }
    if (nav) {
      nav.style.display = expanded ? "" : "none";
      nav.setAttribute("aria-expanded", expanded ? "true" : "false");
    }
  };

  const closeOthers = (currentSection) => {
    for (const section of topLevelSections) {
      if (section === currentSection) {
        continue;
      }
      setExpanded(section, false);
    }
  };

  const activeSection =
    topLevelSections.find((section) =>
      section.classList.contains("md-nav__item--active")
    ) || null;

  for (const section of topLevelSections) {
    const toggle = section.querySelector(":scope > input.md-nav__toggle");
    const label = section.querySelector(":scope > .md-nav__container > label.md-nav__link");
    if (!toggle || !label) {
      continue;
    }

    setExpanded(section, section === activeSection);

    if (!section.dataset.accordionBound) {
      label.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        const shouldExpand = !toggle.checked;
        closeOthers(section);
        setExpanded(section, shouldExpand);
      });
      section.dataset.accordionBound = "true";
    }
  }

    if (!activeSection) {
      closeOthers(null);
    }

    initialized = true;
  };

  const runAccordion = () => {
    initAccordion();
    window.requestAnimationFrame(initAccordion);
    window.setTimeout(initAccordion, 0);
    window.setTimeout(initAccordion, 150);
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", runAccordion, { once: true });
  } else {
    runAccordion();
  }

  window.addEventListener("load", runAccordion);
  window.addEventListener("pageshow", runAccordion);
})();
