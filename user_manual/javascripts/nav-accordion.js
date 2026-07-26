document.addEventListener("DOMContentLoaded", () => {
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

  const closeOthers = (currentSection) => {
    for (const section of topLevelSections) {
      if (section === currentSection) {
        continue;
      }

      const toggle = section.querySelector(":scope > input.md-nav__toggle");
      if (toggle) {
        toggle.checked = false;
      }
    }
  };

  const activeSection =
    topLevelSections.find((section) =>
      section.classList.contains("md-nav__item--active")
    ) || null;

  for (const section of topLevelSections) {
    const toggle = section.querySelector(":scope > input.md-nav__toggle");
    if (!toggle) {
      continue;
    }

    toggle.checked = section === activeSection;
    toggle.addEventListener("change", () => {
      if (toggle.checked) {
        closeOthers(section);
      }
    });

    const nestedLabel = section.querySelector(":scope > .md-nav__link + label.md-nav__link");
    if (nestedLabel) {
      nestedLabel.addEventListener("click", () => {
        window.requestAnimationFrame(() => {
          if (toggle.checked) {
            closeOthers(section);
          }
        });
      });
    }
  }
});
