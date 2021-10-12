import { Dashboard, ItemForm, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";
import { calendarPlus } from "gd-sprest-bs/build/icons/svgs/calendarPlus";
import * as moment from "moment";
import { DataSource, IEventItem } from "./ds";

export class Admin {
  // variables
  private _el: HTMLElement = null;
  private _isAdmin: boolean = false;
  private _dashboard: Dashboard = null;
  private _eventItem: IEventItem = null;
  private _canEditEvent: boolean = false;
  private _canDeleteEvent: boolean = false;

  // Constructor
  constructor(el: HTMLElement, item: IEventItem, dashboard: Dashboard, canEditEvent: boolean, canDeleteEvent: boolean) {
    this._isAdmin = DataSource.IsAdmin;
    this._dashboard = dashboard;
    this._el = el;
    this._eventItem = item;
    this._canEditEvent = canEditEvent;
    this._canDeleteEvent = canDeleteEvent;
    this.renderEventAdminMenu();
  }

  static adminMenuItems(dashboard: Dashboard, canEditEvent: boolean): Components.INavbarItem[] {
    let navItems: Components.INavbarItem[] = [];
    if (DataSource.IsAdmin) {
      navItems.push({
        className: "btn-primary",
        isDisabled: !canEditEvent,
        text: " NEW EVENT",
        isButton: true,
        onClick: () => {
          // Create an item
          ItemForm.create({
            onUpdate: () => {
              // Refresh the dashboard
              //DataSource.refreshDashboard(); 
              // TODO: Fix
            },
            onSetFooter: (elFooter) => {
              let btnGroup = elFooter.querySelector('[role="group"]');
              Components.Button({
                el: btnGroup,
                text: "Cancel",
                type: Components.ButtonTypes.Secondary,
                onClick: () => {
                  ItemForm.close();
                }
              });
              let updateBtn = btnGroup.firstChild as HTMLButtonElement;
              updateBtn.classList.remove("btn-outline-primary");
              updateBtn.classList.add("btn-primary");
            },
            onValidation: (values) => {
              let startDate = values["StartDate"];
              let endDate = values["EndDate"];
              if (moment(startDate).isAfter(moment(endDate))) {
                let ctrl = ItemForm.EditForm.getControl("StartDate");
                ctrl.updateValidation(ctrl.el, {
                  invalidMessage: "Start Date cannot be after the End Date",
                  isValid: false,
                });
                return false;
              }
              return true;
            },
          });
          document.querySelector(".modal-dialog").classList.add("modal-dialog-scrollable");
        },
        iconType: calendarPlus,
        iconSize: 18,
      });
      navItems.push({
        className: "btn-primary",
        isDisabled: !canEditEvent,
        text: "MANAGE GROUPS",
        isButton: true,
        items: [
          {
            text: "Managers",
            href: DataSource.GetManagersUrl(),
            onClick: () => {
              window.open(DataSource.GetManagersUrl(), "_blank");
            },
          },
          {
            text: "Members",
            href: DataSource.GetManagersUrl(),
            onClick: () => {
              window.open(DataSource.GetMembersUrl(), "_blank");
            },
          },
        ],
      });
    }
    return navItems;
  }

  private renderEventAdminMenu() {
    //see if the user if registered
    let isRegistered = this._eventItem.RegisteredUsersId ? this._eventItem.RegisteredUsersId.results.indexOf(ContextInfo.userId) >= 0 : false;
    let adminDropdown = Components.Dropdown({
      el: this._el,
      className: "eventRegAdmin",
      items: [
        {
          isDisabled: !this._canEditEvent,
          text: " Edit",
          onClick: (button) => {
            this.editEvent();
          },
        },
        {
          isDisabled: !this._canDeleteEvent,
          text: " Delete",
          onClick: (button) => {
            this.deleteEvent();
          },
        },
        {
          text: " View Roster",
          onClick: (button) => {
            this.viewRoster();
          },
        },
      ],
    });
    let adminDropdownEl = adminDropdown.el.querySelector("button");
    if (adminDropdownEl) {
      adminDropdownEl.classList.add("btn-icon");
      adminDropdownEl.classList.add("w-100");
      adminDropdownEl.innerHTML =
        '<svg xmlns="http://www.w3.org/2000/svg" width="25" height="25" fill="currentColor" class="bi bi-gear-wide-connected" viewBox="0 0 16 16">' +
        '<path d="M7.068.727c.243-.97 1.62-.97 1.864 0l.071.286a.96.96 0 0 0 1.622.434l.205-.211c.695-.719 1.888-.03 1.613.931l-.08.284a.96.96 0 0 0 1.187 1.187l.283-.081c.96-.275 1.65.918.931 1.613l-.211.205a.96.96 0 0 0 .434 1.622l.286.071c.97.243.97 1.62 0 1.864l-.286.071a.96.96 0 0 0-.434 1.622l.211.205c.719.695.03 1.888-.931 1.613l-.284-.08a.96.96 0 0 0-1.187 1.187l.081.283c.275.96-.918 1.65-1.613.931l-.205-.211a.96.96 0 0 0-1.622.434l-.071.286c-.243.97-1.62.97-1.864 0l-.071-.286a.96.96 0 0 0-1.622-.434l-.205.211c-.695.719-1.888.03-1.613-.931l.08-.284a.96.96 0 0 0-1.186-1.187l-.284.081c-.96.275-1.65-.918-.931-1.613l.211-.205a.96.96 0 0 0-.434-1.622l-.286-.071c-.97-.243-.97-1.62 0-1.864l.286-.071a.96.96 0 0 0 .434-1.622l-.211-.205c-.719-.695-.03-1.888.931-1.613l.284.08a.96.96 0 0 0 1.187-1.186l-.081-.284c-.275-.96.918-1.65 1.613-.931l.205.211a.96.96 0 0 0 1.622-.434l.071-.286zM12.973 8.5H8.25l-2.834 3.779A4.998 4.998 0 0 0 12.973 8.5zm0-1a4.998 4.998 0 0 0-7.557-3.779l2.834 3.78h4.723zM5.048 3.967c-.03.021-.058.043-.087.065l.087-.065zm-.431.355A4.984 4.984 0 0 0 3.002 8c0 1.455.622 2.765 1.615 3.678L7.375 8 4.617 4.322zm.344 7.646.087.065-.087-.065z"/>' +
        "</svg>";
    }
  }

  private editEvent() {
    ItemForm.edit({
      itemId: this._eventItem.Id,
      onUpdate: () => {
        // Refresh the dashboard
        //DataSource.refreshDashboard(); 
        // TODO: Fix
      },
      onSetFooter: (elFooter) => {
        let btnGroup = elFooter.querySelector('[role="group"]');
        Components.Button({
          el: btnGroup,
          text: "Cancel",
          type: Components.ButtonTypes.Secondary,
          onClick: () => {
            ItemForm.close();
          }
        });
        let updateBtn = btnGroup.firstChild as HTMLButtonElement;
        updateBtn.classList.remove("btn-outline-primary");
        updateBtn.classList.add("btn-primary");
      },
      onValidation: (values) => {
        let startDate = values["StartDate"];
        let endDate = values["EndDate"];
        if (moment(startDate).isAfter(moment(endDate))) {
          let ctrl =
            ItemForm.EditForm.getControl("StartDate");
          ctrl.updateValidation(ctrl.el, {
            invalidMessage:
              "Start Date cannot be after the End Date",
            isValid: false,
          });
          return false;
        }
        return true;
      },
    });
    document.querySelector(".modal-dialog").classList.add("modal-dialog-scrollable");
  }

  private deleteEvent() {
    //create the confirmation modal
    let elModal = document.getElementById("event-registration-modal");
    if (elModal == null) {
      //create the element
      elModal = document.createElement("div");
      elModal.className = "modal";
      elModal.id = "event-registration-modal";
      document.body.appendChild(elModal);
    }
    //Create the modal
    let modal = Components.Modal({
      el: elModal,
      title: "Delete Event",
      body: "Are you sure you wanted to delete the selected event?",
      type: Components.ModalTypes.Medium,
      onClose: () => {
        if (elModal) {
          document.body.removeChild(elModal);
          elModal = null;
        }
      },
      onRenderBody: (elBody) => {
        let alert = Components.Alert({
          type: Components.AlertTypes.Danger,
          content: "Error deleting the event",
          className: "eventDeleteErr d-none",
        });
        elBody.prepend(alert.el);
        alert.hide();
      },
      onRenderFooter: (elFooter) => {
        Components.ButtonGroup({
          el: elFooter,
          buttons: [
            {
              text: "Yes",
              type: Components.ButtonTypes.Primary,
              onClick: () => {
                let elAlert =
                  document.querySelector(".eventDeleteErr");
                LoadingDialog.setHeader("Delete Event");
                LoadingDialog.setBody("Deleting the event");
                LoadingDialog.show();
                this._eventItem.delete().execute(
                  () => {
                    // Refresh the dashboard
                    //DataSource.refreshDashboard();
                    // TODO: Fix
                    LoadingDialog.hide();
                    modal.hide();
                  },
                  () => {
                    LoadingDialog.hide();
                    elAlert.classList.remove("d-none");
                  }
                );
              },
            },
            {
              text: "No",
              type: Components.ButtonTypes.Secondary,
              onClick: () => {
                modal.hide();
              },
            },
          ],
        });
      },
    });
    modal.show();
  }

  private viewRoster() {
    //insert the print style css
    let modalCSS =
      "@media print {" +
      "body * {visibility: hidden;}" +
      "#event-registration-modal, #event-registration-modal * {visibility: visible;} " +
      "#event-registration-modal {position: absolute; left: 0; top: 0;} " +
      "div.modal-footer.no-print, div.modal-footer.no-print *, div.modal-header.no-print, div.modal-header.no-print * {display: none;}";
    // Create the style element, if it does not exist
    if (document.getElementById("eventRegPrintStyle") == null) {
      let elStyle = document.createElement("style");
      elStyle.id = "eventRegPrintStyle";
      elStyle.innerHTML = modalCSS;
      document.head.appendChild(elStyle);
    }
    //Get the POCs
    let pocs = this._eventItem["POCId"] ? this._eventItem["POCId"].results : null || [];
    let pocString = "";
    for (let i = 0; i < pocs.length; i++) {
      let poc = Web().getUserById(pocs[i]).executeAndWait().Title;
      if (i > 0) pocString += "<br/>";
      pocString += poc;
    }

    //Get the list of registered usernames
    let usersRegistered = this._eventItem.RegisteredUsersId ? this._eventItem.RegisteredUsersId.results : null || [];
    let usersTable = `<tr><td colspan='4'>No registrations were found for the event</td></tr>`;
    for (let x = 0; x < usersRegistered.length; x++) {
      let userTitle = Web().getUserById(usersRegistered[x]).executeAndWait().Title;
      if (x == 0) usersTable = "";
      usersTable += `<tr><td>${x + 1}</td><td>${userTitle}</td><td></td><td></td></tr>`;
    }

    //create the Jumbotron text
    let jumbotronContent = `<div class='table-responsive' id='print-view'>
                              <table class='table table-striped'>
                                <thead class='thead-dark'>
                                  <tr><th colspan='2' style='text-align:center'><h5>${this._eventItem.Title}</h5></th></tr>
                                </thead>
                                <tbody>
                                  <tr><td><strong>Start Date</strong></td><td> ${moment(this._eventItem.StartDate).format("MM-DD-YYYY HH:mm")}</td></tr>
                                  <tr><td><strong>End Date</strong></td><td>${moment(this._eventItem.EndDate).format("MM-DD-YYYY HH:mm")}</td></tr>
                                  <tr><td><strong>Location</strong></td><td>${this._eventItem.Location}</td></tr>
                                  <tr><td><strong>POC</strong></td><td>${pocString}</td></tr>
                                </tbody>
                              </table>
                              <table class='table table-striped table-hover table-bordered'>
                                <thead class='thead-dark'>
                                  <tr><th colspan='4' style='text-align:center'>Registered Users</th></tr>
                                  <tr>
                                    <th scope='col'>#</th>
                                    <th scope='col'>NAME</th>
                                    <th scope='col'>TIME IN</th>
                                    <th scope='col'>SIGNATURE</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  ${usersTable}
                                </tbody> 
                              </table>
                            </div>`;
    //Create the roster modal
    let elModal = document.getElementById("event-registration-modal");
    if (elModal == null) {
      elModal = document.createElement("div");
      elModal.className = "modal";
      elModal.id = "event-registration-modal";
      document.body.appendChild(elModal);
    }
    let modal = Components.Modal({
      el: elModal,
      id: "rosterModal",
      title: this._eventItem.Title + " - Event Registrations",
      type: Components.ModalTypes.Large,
      onClose: () => {
        if (elModal) {
          document.body.removeChild(elModal);
          elModal = null;
        }
        //Remove the style element for print
        if (document.getElementById("eventRegPrintStyle") != null) {
          document.body.removeChild(document.getElementById("eventRegPrintStyle"));
        }
      },
      onRenderHeader: (elHeader) => {
        elHeader.classList.add("no-print");
      },
      onRenderBody: (elBody) => {
        let jumboTron = Components.Jumbotron({
          el: elBody,
          title: this._eventItem.Title,
          content: jumbotronContent,
        });
      },
      onRenderFooter: (elFooter) => {
        elFooter.classList.add("no-print");
        Components.ButtonGroup({
          el: elFooter,
          buttons: [
            {
              text: "Print",
              type: Components.ButtonTypes.Primary,
              onClick: () => {
                window.print();
              },
            },
            {
              text: "Close",
              type: Components.ButtonTypes.Secondary,
              onClick: () => {
                modal.hide();
              },
            },
          ],
        });
      },
    });
    modal.show();
  }
}