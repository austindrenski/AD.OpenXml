$('.custom-file-input').change(
    function () {

        const fileNames = document.getElementById("file-names");

        for (let child; child = fileNames.firstChild;) {
            fileNames.removeChild(child);
        }

        const files = this.files;

        for (let i = 0; i < files.length; i++) {

            const listItem = document.createElement('li');

            listItem.appendChild(document.createTextNode(files[i].name));

            fileNames.appendChild(listItem);
        }

        $(this).next()
               .after()
               .text(`${files.length} files selected`);
    });
