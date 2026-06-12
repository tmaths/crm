// Main JavaScript file for Report Generator

document.addEventListener('DOMContentLoaded', function() {
    // Auto-hide alerts after 5 seconds
    setTimeout(function() {
        var alerts = document.querySelectorAll('.alert');
        alerts.forEach(function(alert) {
            var bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        });
    }, 3000);

    // Add loading state to buttons when clicked, and always reset after 8 seconds
    document.querySelectorAll('form').forEach(function(form) {
        form.addEventListener('submit', function() {
            var submitButton = form.querySelector('button[type="submit"]');
            if (submitButton) {
                submitButton.disabled = true;
                submitButton.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Processing...';
                // Always reset after 8 seconds
            //     setTimeout(function() {
            //         var original = submitButton.getAttribute('data-original-text');
            //         if (original) {
            //             submitButton.innerHTML = original;
            //         } else {
            //             submitButton.innerHTML = 'Submit';
            //         }
            //         submitButton.disabled = false;
            //     }, 8000);
            }
        });
    });

    // Reset button state on navigation or after file download
    window.addEventListener('pageshow', function() {
        document.querySelectorAll('form').forEach(function(form) {
            var submitButton = form.querySelector('button[type="submit"]');
            if (submitButton && submitButton.disabled) {
                var original = submitButton.getAttribute('data-original-text');
                if (original) {
                    submitButton.innerHTML = original;
                } else {
                    submitButton.innerHTML = 'Submit';
                }
                submitButton.disabled = false;
            }
        });
    });

    // Store original button text for restoration
    document.querySelectorAll('form button[type="submit"]').forEach(function(btn) {
        btn.setAttribute('data-original-text', btn.innerHTML);
    });
}); 