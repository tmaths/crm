function toggleCustomRate(selectElement) {
    const customRateField = document.querySelector('#id_rate');
    const customRateContainer = customRateField.closest('.form-row') || 
                               customRateField.closest('.field-rate') ||
                               customRateField.closest('.field-custom_rate') ||
                               customRateField.parentElement;
    
    if (selectElement.value === 'custom') {
        // Show custom rate field for manual input
        customRateField.required = true;
        customRateField.value = ''; // Clear any preset value
        if (customRateContainer) {
            customRateContainer.classList.remove('hidden');
            customRateContainer.classList.add('show');
        }
        // Focus on the custom rate field
        setTimeout(() => customRateField.focus(), 100);
    } else if (selectElement.value !== '') {
        // Hide custom rate field and set the selected rate
        customRateField.required = false;
        customRateField.value = selectElement.value;
        if (customRateContainer) {
            customRateContainer.classList.add('hidden');
            customRateContainer.classList.remove('show');
        }
    } else {
        // No selection, show empty rate field
        customRateField.required = false;
        customRateField.value = '';
        if (customRateContainer) {
            customRateContainer.classList.remove('hidden');
            customRateContainer.classList.remove('show');
        }
    }
}

// Initialize on page load
document.addEventListener("DOMContentLoaded", function () {
  const rateChoiceField = document.querySelector("#id_rate_choice");
  if (rateChoiceField) {
    // Set initial state
    toggleCustomRate(rateChoiceField);

    // Add event listener
    rateChoiceField.addEventListener("change", function () {
      toggleCustomRate(this);
    });
  }
});
