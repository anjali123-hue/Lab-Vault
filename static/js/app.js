(() => {
    const root = document.documentElement;

    /* =========================================================
       THEME
       ========================================================= */

    const saved = localStorage.getItem("labvault-theme");

    if (saved) {
        root.dataset.theme = saved;
    }

    document.querySelectorAll("[data-theme-toggle]").forEach((button) => {
        button.addEventListener("click", () => {
            const theme =
                root.dataset.theme === "dark"
                    ? "light"
                    : "dark";

            root.dataset.theme = theme;
            localStorage.setItem("labvault-theme", theme);
        });
    });


    /* =========================================================
       PASSWORD SHOW / HIDE
       ========================================================= */

    document.querySelectorAll("[data-password-toggle]").forEach((button) => {
        button.addEventListener("click", () => {
            const input =
                button.parentElement.querySelector("input");

            if (!input) {
                return;
            }

            input.type =
                input.type === "password"
                    ? "text"
                    : "password";

            button.textContent =
                input.type === "password"
                    ? "Show"
                    : "Hide";
        });
    });


    /* =========================================================
       GENERAL MODALS
       ========================================================= */

    document.querySelectorAll("[data-modal-open]").forEach((button) => {
        button.addEventListener("click", () => {
            const modal =
                document.getElementById(button.dataset.modalOpen);

            if (modal) {
                modal.classList.add("open");
                modal.classList.add("is-open");
            }
        });
    });


    document.querySelectorAll("[data-modal-close]").forEach((button) => {
        button.addEventListener("click", () => {
            const modal = button.closest(".modal");

            if (modal) {
                modal.classList.remove("open");
                modal.classList.remove("is-open");
                modal.style.display = "";
            }
        });
    });


    document.querySelectorAll(".modal").forEach((modal) => {
        modal.addEventListener("click", (event) => {
            if (event.target === modal) {
                modal.classList.remove("open");
                modal.classList.remove("is-open");
                modal.style.display = "";
            }
        });
    });


    /* =========================================================
       REJECTION PROMPT
       ========================================================= */

    window.promptReject = (form) => {
        const reason = window.prompt(
            "Why is this request being rejected?"
        );

        if (!reason || !reason.trim()) {
            return false;
        }

        const reasonInput =
            form.querySelector("#reject-reason");

        if (reasonInput) {
            reasonInput.value = reason.trim();
        }

        return true;
    };
})();


/* =========================================================
   STUDENT PRIVACY / CONTACT MODAL
   ========================================================= */

let privacyStudentId = null;


/* ---------------------------------------------------------
   OPEN PRIVACY MODAL
   --------------------------------------------------------- */

function openStudentPrivacyModal(studentId, studentName) {

    privacyStudentId = studentId;

    const modal =
        document.getElementById("student-privacy-modal");

    if (!modal) {
        console.error(
            "Student privacy modal was not found."
        );
        return;
    }

    const nameElement =
        document.getElementById("privacy-student-name");

    const messageElement =
        document.getElementById("privacy-message");

    const passwordSection =
        document.getElementById(
            "privacy-password-section"
        );

    const resultSection =
        document.getElementById(
            "privacy-contact-result"
        );

    const passwordInput =
        document.getElementById(
            "privacy-admin-password"
        );


    if (nameElement) {
        nameElement.textContent =
            studentName || "Student";
    }

    if (messageElement) {
        messageElement.textContent =
            "Enter your Lab Assistant password to view protected contact details.";
    }

    if (passwordSection) {
        passwordSection.style.display = "block";
    }

    if (resultSection) {
        resultSection.style.display = "none";
    }

    if (passwordInput) {
        passwordInput.value = "";
    }


    modal.classList.add("open");
    modal.classList.add("is-open");

    /* Fallback in case another CSS rule is overriding display */
    modal.style.display = "grid";


    setTimeout(() => {
        if (passwordInput) {
            passwordInput.focus();
        }
    }, 100);
}


/* ---------------------------------------------------------
   CLOSE PRIVACY MODAL
   --------------------------------------------------------- */

function closeStudentPrivacyModal() {

    const modal =
        document.getElementById("student-privacy-modal");

    if (modal) {
        modal.classList.remove("open");
        modal.classList.remove("is-open");
        modal.style.display = "";
    }

    privacyStudentId = null;
}


/* ---------------------------------------------------------
   VERIFY ADMIN PASSWORD AND LOAD STUDENT DETAILS
   --------------------------------------------------------- */

async function verifyStudentContact() {

    if (!privacyStudentId) {
        return;
    }


    const passwordInput =
        document.getElementById(
            "privacy-admin-password"
        );

    const messageElement =
        document.getElementById(
            "privacy-message"
        );

    const passwordSection =
        document.getElementById(
            "privacy-password-section"
        );

    const resultSection =
        document.getElementById(
            "privacy-contact-result"
        );


    if (!passwordInput) {
        return;
    }


    const password =
        passwordInput.value.trim();


    if (!password) {

        if (messageElement) {
            messageElement.textContent =
                "Please enter your admin password.";
        }

        passwordInput.focus();

        return;
    }


    if (messageElement) {
        messageElement.textContent =
            "Verifying access...";
    }


    try {

        const response = await fetch(
            `/admin/student/${encodeURIComponent(privacyStudentId)}/contact`,
            {
                method: "POST",

                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                },

                credentials: "same-origin",

                body: JSON.stringify({
                    password: password
                })
            }
        );


        let data;

        try {
            data = await response.json();
        } catch (jsonError) {

            console.error(
                "Invalid JSON response:",
                jsonError
            );

            if (messageElement) {
                messageElement.textContent =
                    "The server returned an invalid response.";
            }

            return;
        }


        if (!response.ok || !data.success) {

            if (messageElement) {
                messageElement.textContent =
                    data.message ||
                    "Unable to verify access.";
            }

            passwordInput.value = "";
            passwordInput.focus();

            return;
        }


        const student =
            data.student || {};


        /* -------------------------------------------------
           FILL STUDENT DETAILS
           ------------------------------------------------- */

        const fields = {
            "privacy-full-name": student.full_name,
            "privacy-uid": student.uid,
            "privacy-email": student.email,
            "privacy-phone": student.phone,
            "privacy-branch": student.branch,
            "privacy-division": student.division,
            "privacy-year": student.year
        };


        Object.entries(fields).forEach(
            ([elementId, value]) => {

                const element =
                    document.getElementById(elementId);

                if (element) {
                    element.textContent =
                        value || "Not provided";
                }

            }
        );


        /* -------------------------------------------------
           SUCCESS MESSAGE
           ------------------------------------------------- */

        if (messageElement) {
            messageElement.textContent =
                "Protected contact information verified successfully.";
        }


        /* -------------------------------------------------
           HIDE PASSWORD FORM
           SHOW STUDENT DETAILS
           ------------------------------------------------- */

        if (passwordSection) {
            passwordSection.style.display = "none";
        }

        if (resultSection) {
            resultSection.style.display = "block";
        }


    } catch (error) {

        console.error(
            "Student contact verification failed:",
            error
        );

        if (messageElement) {
            messageElement.textContent =
                "Unable to verify student information. Please try again.";
        }

    }
}


/* =========================================================
   REQUEST DETAILS MODAL
   ========================================================= */

let requestDetailsId = null;


/* ---------------------------------------------------------
   OPEN REQUEST DETAILS MODAL
   --------------------------------------------------------- */

function openRequestDetailsModal(requestId, requestCode) {

    requestDetailsId = requestId;

    const modal =
        document.getElementById("request-details-modal");

    if (!modal) {
        console.error(
            "Request details modal was not found."
        );
        return;
    }

    const titleElement =
        document.getElementById(
            "request-details-title"
        );

    const messageElement =
        document.getElementById(
            "request-details-message"
        );

    const passwordSection =
        document.getElementById(
            "request-details-password-section"
        );

    const resultSection =
        document.getElementById(
            "request-details-result"
        );

    const passwordInput =
        document.getElementById(
            "request-details-password"
        );


    if (titleElement) {
        titleElement.textContent =
            requestCode || "Request details";
    }

    if (messageElement) {
        messageElement.textContent =
            "Enter your Admin password to view protected student and component details.";
    }

    if (passwordInput) {
        passwordInput.value = "";
    }

    if (passwordSection) {
        passwordSection.style.display = "block";
    }

    if (resultSection) {
        resultSection.style.display = "none";
    }


    /* ---------------------------------------------------------
       OPEN MODAL
       --------------------------------------------------------- */

    modal.classList.add("open");
    modal.classList.add("is-open");

    /*
       Your CSS already contains:
       .modal.is-open { display:flex; }

       We also set display directly as a fallback so the modal
       cannot remain hidden because of another CSS rule.
    */
    modal.style.display = "flex";


    setTimeout(() => {
        if (passwordInput) {
            passwordInput.focus();
        }
    }, 100);
}


/* ---------------------------------------------------------
   CLOSE REQUEST DETAILS MODAL
   --------------------------------------------------------- */

function closeRequestDetailsModal() {

    const modal =
        document.getElementById(
            "request-details-modal"
        );

    if (modal) {
        modal.classList.remove("open");
        modal.classList.remove("is-open");

        /* Remove fallback inline display */
        modal.style.display = "";
    }

    requestDetailsId = null;
}


/* ---------------------------------------------------------
   VERIFY ADMIN PASSWORD AND LOAD REQUEST DETAILS
   --------------------------------------------------------- */

async function verifyRequestDetails() {

    if (!requestDetailsId) {
        return;
    }


    const passwordInput =
        document.getElementById(
            "request-details-password"
        );

    const messageElement =
        document.getElementById(
            "request-details-message"
        );

    const passwordSection =
        document.getElementById(
            "request-details-password-section"
        );

    const resultSection =
        document.getElementById(
            "request-details-result"
        );


    if (!passwordInput) {
        return;
    }


    const password =
        passwordInput.value.trim();


    if (!password) {

        if (messageElement) {
            messageElement.textContent =
                "Please enter your Admin password.";
        }

        passwordInput.focus();

        return;
    }


    if (messageElement) {
        messageElement.textContent =
            "Verifying access...";
    }


    try {

        const response = await fetch(
            `/admin/request/${encodeURIComponent(requestDetailsId)}/details`,
            {
                method: "POST",

                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                },

                credentials: "same-origin",

                body: JSON.stringify({
                    password: password
                })
            }
        );


        let data;

        try {
            data = await response.json();
        } catch (jsonError) {

            console.error(
                "Invalid JSON response:",
                jsonError
            );

            if (messageElement) {
                messageElement.textContent =
                    "The server returned an invalid response.";
            }

            return;
        }


        if (!response.ok || !data.success) {

            if (messageElement) {
                messageElement.textContent =
                    data.message ||
                    "Unable to verify access.";
            }

            passwordInput.value = "";
            passwordInput.focus();

            return;
        }


        const req =
            data.request || {};

        const items =
            Array.isArray(data.items)
                ? data.items
                : [];


        /* ---------------------------------------------------------
           REQUEST INFORMATION
           --------------------------------------------------------- */

        const studentElement =
            document.getElementById(
                "request-detail-student"
            );

        const erpElement =
            document.getElementById(
                "request-detail-erp"
            );

        const uidElement =
            document.getElementById(
                "request-detail-uid"
            );

        const phoneElement =
            document.getElementById(
                "request-detail-phone"
            );

        const emailElement =
            document.getElementById(
                "request-detail-email"
            );

        const branchElement =
            document.getElementById(
                "request-detail-branch"
            );

        const yearElement =
            document.getElementById(
                "request-detail-year"
            );

        const statusElement =
            document.getElementById(
                "request-detail-status"
            );

        const projectElement =
            document.getElementById(
                "request-detail-project"
            );

        const purposeElement =
            document.getElementById(
                "request-detail-purpose"
            );

        const dueElement =
            document.getElementById(
                "request-detail-due"
            );


        if (studentElement) {
            studentElement.textContent =
                req.student_name || "Not provided";
        }

        if (erpElement) {
            erpElement.textContent =
                req.erp_id || "Not provided";
        }

        if (uidElement) {
            uidElement.textContent =
                req.uid || "Not provided";
        }

        if (phoneElement) {
            phoneElement.textContent =
                req.phone || "Not provided";
        }

        if (emailElement) {
            emailElement.textContent =
                req.email || "Not provided";
        }

        if (branchElement) {
            branchElement.textContent =
                `${req.branch || "Not provided"} / ${req.division || "Not provided"}`;
        }

        if (yearElement) {
            yearElement.textContent =
                req.year || "Not provided";
        }

        if (statusElement) {

            const status =
                req.overall_status || "";

            statusElement.textContent =
                status
                    .replaceAll("_", " ")
                    .replace(/\b\w/g, (c) => c.toUpperCase());
        }

        if (projectElement) {
            projectElement.textContent =
                req.project_title || "Not provided";
        }

        if (purposeElement) {
            purposeElement.textContent =
                req.purpose || "Not provided";
        }

        if (dueElement) {
            dueElement.textContent =
                req.due_date || "Not provided";
        }


        /* ---------------------------------------------------------
           COMPONENTS
           --------------------------------------------------------- */

        const componentContainer =
            document.getElementById(
                "request-detail-components"
            );


        if (componentContainer) {

            componentContainer.innerHTML = "";


            if (!items.length) {

                const empty =
                    document.createElement("div");

                empty.className =
                    "empty-small";

                empty.innerHTML = `
                    <strong>No component records found.</strong>
                    <span>
                        This request does not currently have
                        component entries.
                    </span>
                `;

                componentContainer.appendChild(
                    empty
                );

            } else {

                items.forEach((item) => {

                    const chip =
                        document.createElement("span");

                    chip.className =
                        "pending-chip";

                    chip.innerHTML = `
                        ${escapeHtml(item.name || "Unknown component")}
                        <b>× ${item.quantity || 0}</b>
                    `;

                    componentContainer.appendChild(
                        chip
                    );

                });
            }
        }


        /* ---------------------------------------------------------
           SUCCESS
           --------------------------------------------------------- */

        if (messageElement) {
            messageElement.textContent =
                "Protected request information verified successfully.";
        }

        if (passwordSection) {
            passwordSection.style.display =
                "none";
        }

        if (resultSection) {
            resultSection.style.display =
                "block";
        }

        /*
           Keep the modal open after successful verification.
        */
        const modal =
            document.getElementById(
                "request-details-modal"
            );

        if (modal) {
            modal.classList.add("open");
            modal.classList.add("is-open");
            modal.style.display = "flex";
        }


    } catch (error) {

        console.error(
            "Request details verification failed:",
            error
        );

        if (messageElement) {
            messageElement.textContent =
                "Unable to load request details. Please try again.";
        }
    }
}


/* =========================================================
   ESCAPE HTML
   ========================================================= */

function escapeHtml(value) {

    const div =
        document.createElement("div");

    div.textContent =
        value ?? "";

    return div.innerHTML;
}


/* =========================================================
   ESC KEY FOR MODALS
   ========================================================= */

document.addEventListener("keydown", (event) => {

    if (event.key !== "Escape") {
        return;
    }


    const privacyModal =
        document.getElementById(
            "student-privacy-modal"
        );

    if (
        privacyModal &&
        (
            privacyModal.classList.contains("open") ||
            privacyModal.classList.contains("is-open")
        )
    ) {
        closeStudentPrivacyModal();
        return;
    }


    const requestModal =
        document.getElementById(
            "request-details-modal"
        );

    if (
        requestModal &&
        (
            requestModal.classList.contains("open") ||
            requestModal.classList.contains("is-open")
        )
    ) {
        closeRequestDetailsModal();
    }

});

(function () {
  const miniBar = document.querySelector(".mini-footer-bar");
  const fullFooter = document.querySelector(".site-footer");

  if (!miniBar || !fullFooter) return;

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        miniBar.style.transform = entry.isIntersecting
          ? "translateY(100%)"
          : "translateY(0)";
      });
    },
    { threshold: 0 }
  );

  observer.observe(fullFooter);
})();