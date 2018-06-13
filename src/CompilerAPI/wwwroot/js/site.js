$(".custom-file-input").change(function () {
    const n = document.getElementById("file-names");
    for (let t; t = n.firstChild;) n.removeChild(t);
    const t = this.files;
    for (let i = 0; i < t.length; i++) {
        const r = document.createElement("li");
        r.appendChild(document.createTextNode(t[i].name));
        n.appendChild(r)
    }
    $(this).next().after().text(`${t.length} files selected`)
});