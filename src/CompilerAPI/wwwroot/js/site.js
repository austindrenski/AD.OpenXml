$('.custom-file').change(customFileInputLabel);
$('#file-upload').change(fileUploadList);

function customFileInputLabel() {
    const files = $(this).find('.custom-file-input').first().prop('files');
    const label = $(this).find('.custom-file-label').first();

    switch (files.length) {
        case 0: {
            label.text(label.attr('placeholder'));
            return;
        }
        case 1: {
            label.text(files[0].name);
            return;
        }
        default: {
            label.text(`${files.length} files selected`);
            return;
        }
    }
}

function fileUploadList() {
    const names = $('#file-names').first();
    const files = $('#files').prop('files');

    names.empty();

    for (let i = 0; i < files.length; i++) {
        names.append(`<li>${files[i].name}</li>`);
    }
}