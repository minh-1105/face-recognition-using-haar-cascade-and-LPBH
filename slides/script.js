const slides = Array.from(document.querySelectorAll(".slide"));
const progressBar = document.getElementById("progress-bar");
const slideCounter = document.getElementById("slide-counter");
let currentIndex = 0;

function renderSlide(index) {
  currentIndex = Math.max(0, Math.min(index, slides.length - 1));

  slides.forEach((slide, slideIndex) => {
    slide.classList.toggle("active", slideIndex === currentIndex);
  });

  const progress = ((currentIndex + 1) / slides.length) * 100;
  progressBar.style.width = `${progress}%`;
  slideCounter.textContent = `${currentIndex + 1} / ${slides.length}`;
  document.title = `Haar Cascade Deck · ${currentIndex + 1}/${slides.length}`;
}

function nextSlide() {
  renderSlide(currentIndex + 1);
}

function previousSlide() {
  renderSlide(currentIndex - 1);
}

document.addEventListener("keydown", (event) => {
  if (["ArrowRight", " ", "Enter"].includes(event.key)) {
    event.preventDefault();
    nextSlide();
  }

  if (event.key === "ArrowLeft") {
    event.preventDefault();
    previousSlide();
  }

  if (event.key === "Home") {
    renderSlide(0);
  }

  if (event.key === "End") {
    renderSlide(slides.length - 1);
  }
});

document.addEventListener("click", (event) => {
  const viewportMidpoint = window.innerWidth / 2;
  if (event.clientX >= viewportMidpoint) {
    nextSlide();
    return;
  }

  previousSlide();
});

renderSlide(0);
