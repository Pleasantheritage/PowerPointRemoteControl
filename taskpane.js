Office.onReady(() => {
    document.getElementById("nextSlide").addEventListener("click", () => controlSlides("NEXT"));
    document.getElementById("prevSlide").addEventListener("click", () => controlSlides("PREV"));
    document.getElementById("startPresentation").addEventListener("click", () => controlSlides("START"));
    document.getElementById("endPresentation").addEventListener("click", () => controlSlides("END"));
    document.getElementById("goToSlideBtn").addEventListener("click", () => controlSlides("GOTO"));
});

// Function to control slides
function controlSlides(command) {
    switch (command) {
        case "NEXT":
            Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index);
            break;
        case "PREV":
            Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index);
            break;
        case "START":
            Office.context.document.goToByIdAsync(1, Office.GoToType.Slide); // First slide
            break;
        case "END":
            Office.context.document.getFileAsync(Office.FileType.Pdf, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    let numSlides = result.value.length;
                    Office.context.document.goToByIdAsync(numSlides, Office.GoToType.Slide); // Last slide
                }
            });
            break;
        case "GOTO":
            goToSlide();
            break;
    }
}

// Function to go to a specific slide number
function goToSlide() {
    let slideNum = parseInt(document.getElementById("slideNumber").value);

    if (isNaN(slideNum) || slideNum < 1) {
        alert("Please enter a valid slide number.");
        return;
    }

    Office.context.document.goToByIdAsync(slideNum, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Error going to slide:", asyncResult.error.message);
            alert("Error: " + asyncResult.error.message);
        }
    });
}
