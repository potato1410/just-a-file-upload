const canvas = document.getElementById("gameCanvas");  
const ctx = canvas.getContext("2d");  const birdImg = new Image();
birdImg.src = "Musk.jpg";

let birdY = 0;
let gravity = 0;
let jump = -0.5;
let velocity = 0.1;

birdImg.onload = function () {
requestAnimationFrame(draw);
};

function draw() {
ctx.clearRect(0, 0, canvas.width, canvas.height);

// Draw the bird at the current Y position
ctx.drawImage(birdImg, 100, birdY, 50, 50);

gravity += velocity;
birdY += gravity;
document.addEventListener("keydown", function(event) {
if (event.code === "Space") {
velocity = -0.5;

setTimeout(function() {  
    velocity = 0.1;  
    output.textContent = "velocity: " + velocity;  
  }, 100); // 1000 Millisekunden = 1 Sekunde  
}

});

// Stop after it goes below the canvas
if (birdY < canvas.height - 50) {
requestAnimationFrame(draw);
}
}
