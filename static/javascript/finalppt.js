const slideContainer = document.querySelector('.slide-container');
const numSlides = 6; // Number of slides

// Function to create slides dynamically
function createSlides() {
    for (let i = 1; i <= numSlides; i++) {
        const slide = document.createElement('div');
        slide.classList.add('slide');
        slide.innerHTML = `
            <div><img src="../bg${i}.png" alt="Slide ${i}"></div>
        `;
        slideContainer.appendChild(slide);
    }
}

// Create the slides on initial load
createSlides();

let currentIndex = 0;

// Function to update slides to show the current index
function updateSlides() {
    const slides = slideContainer.children;
    const targetSlide = slides[currentIndex];
    targetSlide.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Keydown event listener for navigation
document.addEventListener('keydown', (e) => {
    const slides = slideContainer.children.length;
    if (e.key === 'ArrowDown') {
        currentIndex = (currentIndex + 1) % slides; // Move down
        updateSlides();
    } else if (e.key === 'ArrowUp') {
        currentIndex = (currentIndex - 1 + slides) % slides; // Move up
        updateSlides();
    }
});
