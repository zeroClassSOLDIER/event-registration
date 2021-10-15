import { CanvasForm, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, List, Types } from "gd-sprest-bs";
import { IEventItem } from "./ds";
import { fileEarmarkText } from "gd-sprest-bs/build/icons/svgs/fileEarmarkText";
import { fileEarmarkX } from "gd-sprest-bs/build/icons/svgs/fileEarmarkX";
import { files } from "gd-sprest-bs/build/icons/svgs/files";
import { upload } from "gd-sprest-bs/build/icons/svgs/upload";
import Strings from "./strings"

/**
 * Documents View
 */
export class DocumentsView {
    // private variables
    private _item: IEventItem = null;
    private _el: HTMLElement = null;
    private _isAdmin: boolean = false;
    private _canEditEvent: boolean = false;
    private _onRefresh: () => void = null;

    // Constructor
    constructor(el: HTMLElement, item: IEventItem, isAdmin: boolean, canEditEvent: boolean, onRefresh: () => void) {
        this._item = item;
        this._el = el;
        this._isAdmin = isAdmin;
        this._canEditEvent = canEditEvent;
        this._onRefresh = onRefresh;
        this.render();
    }

    // Display the delete modal
    private deleteDocument(attachment: Types.SP.Attachment) {
        // Hide the slider
        CanvasForm.hide();

        // Set the header
        Modal.setHeader("Delete " + attachment.FileName);

        // Set the body
        Modal.setBody("Are you sure you want to delete the document?");

        // Hide the auto close
        Modal.setAutoClose(false);
        Modal.setCloseEvent(() => {
            // Show the slider
            CanvasForm.show();
        });

        // Set the footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Yes",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Close the dialog
                        Modal.hide();

                        // Show a loading dialog
                        LoadingDialog.setHeader("Deleting Document");
                        LoadingDialog.setBody("This will close after the document has been deleted.");

                        // Delete the attachment
                        attachment.delete().execute(
                            // Success
                            () => {
                                // Parse the attachments
                                for (let i = 0; i < this._item.AttachmentFiles.results.length; i++) {
                                    let filename = this._item.AttachmentFiles.results[i].FileName;

                                    // See if this is the target item
                                    if (attachment.FileName == filename) {
                                        // Remove this item
                                        this._item.AttachmentFiles.results.splice(i, 1);
                                        break;
                                    }
                                }

                                // View the attachments
                                this.viewAttachments();

                                // Hide the loading dialog
                                LoadingDialog.hide();
                            },

                            // Error
                            () => {
                                // TODO
                                // Hide the loading dialog
                                LoadingDialog.hide();
                            }
                        )
                    }
                },
                {
                    text: "No",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => {
                        // Close the dialog
                        Modal.hide();

                        // Show the slider
                        CanvasForm.show();
                    }
                }
            ]
        }).el);

        // Show the modal
        Modal.show();
    }

    // Renders the attachment icons
    private render() {
        // See if attachments exist
        if (this._item.AttachmentFiles.results.length > 0) {
            // Render the view attachments tooltip
            Components.Tooltip({
                el: this._el,
                content: "View Attachments",
                placement: Components.TooltipPlacements.Top,
                btnProps: {
                    className: "me-2",
                    iconType: files,
                    iconSize: 24,
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Clear the canvas form and modal
                        CanvasForm.clear();
                        Modal.clear();

                        // View the attachments
                        this.viewAttachments();
                    }
                }
            });
        }

        //Render the upload tooltip
        Components.Tooltip({
            el: this._el,
            content: "Upload a document",
            placement: Components.TooltipPlacements.Top,
            btnProps: {
                isDisabled: !this._canEditEvent,
                iconType: upload,
                iconSize: 24,
                toggle: "tooltip",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Show the upload document dialog
                    this.uploadDocument();
                },
            }
        });
    }

    // Uploads a new attachment file
    private uploadDocument() {
        // Show the file upload dialog
        Helper.ListForm.showFileDialog().then((fileInfo) => {
            // Set the loading dialog
            LoadingDialog.setHeader("Upload Document");
            LoadingDialog.setBody("Attaching the document to the event.");
            LoadingDialog.show();

            // Add the attachment
            List(Strings.Lists.Events)
                .Items(this._item.Id)
                .AttachmentFiles()
                .add(fileInfo.name, fileInfo.data)
                .execute(
                    // Success
                    () => {
                        // Refresh the dashboard
                        this._onRefresh();

                        // Hide the dialog
                        LoadingDialog.hide();
                    },

                    // Error
                    (err) => {
                        // TODO
                        // Log to the console
                        console.log("Error uploading the file", err);

                        // Hide the dialog
                        LoadingDialog.hide();
                    }
                );
        });
    }

    // Displays the attachments in a slider
    private viewAttachments() {
        // Set the header
        CanvasForm.setHeader("View Attachments");

        // Set the body
        CanvasForm.setBody(Components.Table({
            rows: this._item.AttachmentFiles.results,
            columns: [
                {
                    name: "FileName",
                    title: "File Name"
                },
                {
                    name: "",
                    onRenderCell: (el, col, attachment: Types.SP.Attachment) => {
                        // Render the view tooltip
                        Components.Tooltip({
                            el,
                            content: "View the document",
                            btnProps: {
                                iconType: fileEarmarkText,
                                iconSize: 24,
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    let isWopi: boolean = Documents.isWopi(attachment.FileName);
                                    window.open(
                                        isWopi
                                            ? ContextInfo.webServerRelativeUrl +
                                            "/_layouts/15/WopiFrame.aspx?sourcedoc=" +
                                            attachment.ServerRelativeUrl +
                                            "&action=view"
                                            : attachment.ServerRelativeUrl,
                                        "_blank"
                                    );
                                }
                            }
                        });

                        // See if this is an admin
                        if (this._isAdmin) {
                            // Render the delete tooltip
                            Components.Tooltip({
                                el,
                                content: "Delete the document",
                                btnProps: {
                                    className: "ms-2",
                                    isDisabled: !this._canEditEvent,
                                    iconType: fileEarmarkX,
                                    iconSize: 24,
                                    toggle: "tooltip",
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Delete the document
                                        this.deleteDocument(attachment);
                                    }
                                }
                            });
                        }
                    }
                }
            ]
        }).el);

        // Show the slider
        CanvasForm.show();
    }
}