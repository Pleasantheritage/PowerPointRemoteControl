Office.onReady(() => {
    document.getElementById("nextSlide").addEventListener("click", () => controlSlides("NEXT"));
    document.getElementById("prevSlide").addEventListener("click", () => controlSlides("PREV"));
    document.getElementById("startPresentation").addEventListener("click", () => controlSlides("START"));
    document.getElementById("endPresentation").addEventListener("click", () => controlSlides("END"));
});

function controlSlides(command) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            Office.context.document.addHandlerAsync(
                Office.EventType.DocumentSelectionChanged, 
                () => executeCommand(command)
            );
        }
    });
}

function executeCommand(command) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            switch (command) {
                case "NEXT": Office.context.document.goToByIdAsync(Office.GoToType.Slide, { index: 1 }); break;
                case "PREV": Office.context.document.goToByIdAsync(Office.GoToType.Slide, { index: -1 }); break;
                case "START": Office.context.document.goToByIdAsync(Office.GoToType.Slide, { index: 0 }); break;
                case "END": Office.context.document.goToByIdAsync(Office.GoToType.Slide, { index: 999 }); break;
            }
        }
    });
}
