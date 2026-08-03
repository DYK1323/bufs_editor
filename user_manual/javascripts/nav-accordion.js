(() => {
  const normalizeChapterFiveLinks = () => {
    const sidebar = document.querySelector(".md-sidebar--primary");
    if (!sidebar) {
      return;
    }

    const homeLink =
      document.querySelector(".md-header__button.md-logo") ||
      document.querySelector(".md-nav__button.md-logo");
    const docsRoot = homeLink
      ? new URL(homeLink.getAttribute("href") || ".", window.location.href)
      : new URL("/", window.location.href);

    const chapterFiveTargets = new Map([
      ["5.5 문단 탭", "05/05-paragraph-tab/"],
      ["5.6 숫자와 날짜 정리", "05/05-05/"],
      ["5.7 표 정리", "05/05-06/"],
      ["5.8 캡션 만들기", "05/05-07/"],
      ["5.9 마크다운 표 변환", "05/05-08/"],
    ]);

    for (const link of sidebar.querySelectorAll("a.md-nav__link")) {
      const text = (link.textContent || "").replace(/\s+/g, " ").trim();
      const target = chapterFiveTargets.get(text);
      if (!target) {
        continue;
      }
      link.href = new URL(target, docsRoot).href;
    }
  };

  const initAccordion = () => {
    normalizeChapterFiveLinks();

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

    const closeAll = () => {
      for (const section of topLevelSections) {
        setExpanded(section, false);
      }
    };

    const activeSection =
      topLevelSections.find((section) =>
        section.classList.contains("md-nav__item--active")
      ) || null;

    for (const section of topLevelSections) {
      const toggle = section.querySelector(":scope > input.md-nav__toggle");
      const container = section.querySelector(":scope > .md-nav__container");
      const label = container?.querySelector("label.md-nav__link");
      const link = container?.querySelector("a.md-nav__link");
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

        if (link) {
          link.addEventListener("click", (event) => {
            if (
              event.defaultPrevented ||
              event.button !== 0 ||
              event.metaKey ||
              event.ctrlKey ||
              event.shiftKey ||
              event.altKey
            ) {
              return;
            }
            closeAll();
          });
        }

        section.dataset.accordionBound = "true";
      }
    }

    if (!activeSection) {
      closeOthers(null);
    }
  };

  const runAccordion = () => {
    normalizeChapterFiveLinks();
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
