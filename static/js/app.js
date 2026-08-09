(() => {
  const root = document.documentElement;
  const saved = localStorage.getItem("labvault-theme");
  if (saved) root.dataset.theme = saved;
  document.querySelectorAll("[data-theme-toggle]").forEach((button) => {
    button.addEventListener("click", () => {
      const theme = root.dataset.theme === "dark" ? "light" : "dark";
      root.dataset.theme = theme;
      localStorage.setItem("labvault-theme", theme);
    });
  });
  document.querySelectorAll("[data-password-toggle]").forEach((button) => {
    button.addEventListener("click", () => {
      const input = button.parentElement.querySelector("input");
      input.type = input.type === "password" ? "text" : "password";
      button.textContent = input.type === "password" ? "Show" : "Hide";
    });
  });
  document.querySelectorAll("[data-modal-open]").forEach((button) => {
    button.addEventListener("click", () => document.getElementById(button.dataset.modalOpen)?.classList.add("open"));
  });
  document.querySelectorAll("[data-modal-close]").forEach((button) => {
    button.addEventListener("click", () => button.closest(".modal")?.classList.remove("open"));
  });
  document.querySelectorAll(".modal").forEach((modal) => {
    modal.addEventListener("click", (event) => { if (event.target === modal) modal.classList.remove("open"); });
  });
  window.promptReject = (form) => {
    const reason = window.prompt("Why is this request being rejected?");
    if (!reason || !reason.trim()) return false;
    form.querySelector("#reject-reason").value = reason.trim();
    return true;
  };
})();