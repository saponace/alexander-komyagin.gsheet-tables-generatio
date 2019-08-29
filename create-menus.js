function onOpen(e) {
    DocumentApp.getUi()
        .createMenu('Run scripts')
        .addItem('Generate document', 'generateDoc')
        .addToUi();
}
