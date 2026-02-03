(() => {
  // Theme toggle
  const btn = document.getElementById("btnTheme");
  if (btn) {
    btn.addEventListener("click", () => {
      const cur = document.documentElement.getAttribute("data-theme") || "dark";
      const next = cur === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", next);
      localStorage.setItem("theme", next);
    });
  }

  // Auto-show bootstrap toasts
  document.querySelectorAll(".toast").forEach(t => {
    const toast = new bootstrap.Toast(t, { delay: 3500 });
    toast.show();
  });
})();
