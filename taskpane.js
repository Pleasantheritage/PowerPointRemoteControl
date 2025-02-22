let socket;
const SERVER_URL = 'ws://localhost:8080'; // Address of your Windows BLE server

// Connect to the WebSocket server
function connectToWebSocket() {
  socket = new WebSocket(SERVER_URL);

  socket.onopen = function () {
    console.log("Connected to WebSocket server");
  };

  socket.onmessage = function (event) {
    const message = JSON.parse(event.data);
    handleBluetoothMessage(message);
  };

  socket.onerror = function (error) {
    console.error("WebSocket Error: ", error);
  };

  socket.onclose = function () {
    console.log("Disconnected from WebSocket server");
  };
}

// Handle the Bluetooth message and control slides accordingly
function handleBluetoothMessage(message) {
  if (message.command === "NEXT") {
    controlSlides("NEXT");
  } else if (message.command === "PREV") {
    controlSlides("PREV");
  } else if (message.command === "START") {
    controlSlides("START");
  } else if (message.command === "END") {
    controlSlides("END");
  }
}

// Send slide control commands via WebSocket
function controlSlides(command) {
  Office.context.document.goToByIdAsync(Office.GoToType.Slide, {
    index: command === "NEXT" ? 1 : command === "PREV" ? -1 : command === "START" ? 0 : 999
  });
}

// Function to go to a specific slide number
function goToSlideNumber(slideNumber) {
  Office.context.document.goToByIdAsync(Office.GoToType.Slide, {
    index: slideNumber - 1 // Zero-based index for slides
  });
}

// Initialize WebSocket connection
connectToWebSocket();

// Add listeners for buttons in the taskpane (e.g., for manual control via buttons)
Office.onReady(() => {
  document.getElementById("nextSlide").addEventListener("click", () => controlSlides("NEXT"));
  document.getElementById("prevSlide").addEventListener("click", () => controlSlides("PREV"));
  document.getElementById("startPresentation").addEventListener("click", () => controlSlides("START"));
  document.getElementById("endPresentation").addEventListener("click", () => controlSlides("END"));
  
  // Add event listener for "Go to Slide" button
  document.getElementById("goToSlide").addEventListener("click", () => {
    const slideNumber = parseInt(document.getElementById("slideNumberInput").value);
    if (!isNaN(slideNumber) && slideNumber > 0) {
      goToSlideNumber(slideNumber);
    } else {
      alert("Please enter a valid slide number.");
    }
  });
});
