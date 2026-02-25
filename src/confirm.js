Office.onReady(() => {

    Office.actions.associate("openConfirmDialog", function (event) {

        const dialogUrl =
            "https://dellasiegaexternal2.github.io/Herisson24022026/src/confirm.html?mode=dialog&v=100";

        Office.context.ui.displayDialogAsync(
            dialogUrl,
            {
                height: 70,
                width: 60,
                displayInIframe: true
            },
            function (asyncResult) {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error(asyncResult.error.message);
                }
            }
        );

        event.completed();
    });

});
