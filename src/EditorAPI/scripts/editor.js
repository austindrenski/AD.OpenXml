"use strict";
require("monaco-editor/monaco");
var Editor = monaco.editor.create(document.getElementById('editor'), {
    value: '# Welcome to the Markdown editor demo!',
    language: 'markdown',
    minimap: { enabled: false }
});
document.getElementById('undoButton').onclick = Undo;
document.getElementById('redoButton').onclick = Redo;
Editor.focus();
var InitialVersion = Editor.getModel().getAlternativeVersionId();
var CurrentVersion = InitialVersion;
var LastVersion = InitialVersion;
Editor.onDidChangeModelContent(function (e) {
    var VersionId = Editor.getModel().getAlternativeVersionId();
    // undoing
    if (VersionId < CurrentVersion) {
        EnableRedoButton();
        // no more undo possible
        if (VersionId === InitialVersion) {
            DisableUndoButton();
        }
    }
    else {
        // redoing
        if (VersionId <= LastVersion) {
            // redoing the last change
            if (VersionId === LastVersion) {
                DisableRedoButton();
            }
        }
        else { // adding new change, disable redo when adding new changes
            DisableRedoButton();
            if (CurrentVersion > LastVersion) {
                LastVersion = CurrentVersion;
            }
        }
        EnableUndoButton();
    }
    CurrentVersion = VersionId;
});
function Undo() {
    Editor.trigger('undo', 'undo', null);
    Editor.focus();
}
function Redo() {
    Editor.trigger('redo', 'redo', null);
    Editor.focus();
}
function EnableUndoButton() {
    document.getElementById("undoButton").disabled = false;
}
function DisableUndoButton() {
    document.getElementById("undoButton").disabled = true;
}
function EnableRedoButton() {
    document.getElementById("redoButton").disabled = false;
}
function DisableRedoButton() {
    document.getElementById("redoButton").disabled = true;
}
module.exports = 0;
//# sourceMappingURL=editor.js.map