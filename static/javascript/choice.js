document.addEventListener("DOMContentLoaded", function() {
    const closeModalButton = document.querySelector('.close-btn');
    const stylesContainer = document.getElementById('styles-container');
    const generateButton = document.querySelector('.generate-btn');
    const form = document.getElementById('presentation-form');
    const slidesInput = document.getElementById('slides');
    const slideError = document.getElementById('slide-error');

    let numberOfSlides = slidesInput.value;
    let selectedStyle = '';
    let selectedPPTPath = '';  // Variable to store the path of the selected template

    // Standard template list
    const templates = [
        { src: '/static/images/bg1.png', alt: 'Minimalist', label: 'Minimalist' },
        { src: '/static/images/bg2.png', alt: 'Geometric', label: 'Geometric' },
        { src: '/static/images/bg3.jpg', alt: 'Colorful', label: 'Colorful' },
        { src: '/static/images/bg8.jpeg', alt: 'Professional', label: 'Professional' },
        // { src: '/static/images/royal.jpg', alt: 'Royal Company Profile', label: 'Royal Company Profile' }
    ];

    // Function to populate styles with standard templates
    function updateStyles() {
        stylesContainer.innerHTML = ''; // Clear existing styles

        templates.forEach((template, index) => {
            const button = document.createElement('button');
            button.classList.add('style-btn');

            const img = document.createElement('img');
            img.src = template.src;
            img.alt = template.alt;

            const p = document.createElement('p');
            p.textContent = template.label;

            button.appendChild(img);
            button.appendChild(p);
            stylesContainer.appendChild(button);

            // Automatically select the first template
            if (index === 0) {
                button.classList.add('active');
                selectedStyle = template.label;  // Store the first template as the default
                selectedPPTPath = template.src;  // Store the path of the first template
            }

            // Add event listener for each style button
            button.addEventListener('click', function(event) {
                event.preventDefault();  // Prevent default behavior (redirecting)

                document.querySelectorAll('.style-btn').forEach(btn => btn.classList.remove('active'));
                button.classList.add('active');
                selectedStyle = template.label;  // Update the selected style
                selectedPPTPath = template.src;  // Store the path of the selected template
            });
        });
    }

    // Initial population of standard templates
    updateStyles();

    // Event listener to close the modal and go to the previous page
    closeModalButton.addEventListener('click', function() {
        window.history.back(); // Go to the previous page
    });

    // Validate the number of slides input
    slidesInput.addEventListener('input', function() {
        numberOfSlides = parseInt(slidesInput.value, 10);

        if (numberOfSlides < 6) {
            slideError.style.display = 'block';
            generateButton.disabled = true;
        } else {
            slideError.style.display = 'none';
            generateButton.disabled = false;
        }
    });

    // Event listener for form submission
    form.addEventListener('submit', function(event) {
        event.preventDefault();
        
        numberOfSlides = slidesInput.value;

        // Store the data in variables
        const formData = {
            numberOfSlides: numberOfSlides,
            selectedStyle: selectedStyle,
            pptPath: selectedPPTPath  // This stores the selected template's image path
        };

        // Log the formData to the console
        console.log(formData);

        // Redirect to the next page
        window.location.href = 'presentation.html'; // Change 'nextpage.html' to your actual target page
    });
});
