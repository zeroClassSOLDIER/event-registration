import { InstallationRequired, ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components, Web } from "gd-sprest-bs";
import { calendarPlus } from "gd-sprest-bs/build/icons/svgs/calendarPlus";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs";
import * as moment from "moment";
import { DataSource, IEventItem } from "./ds";

export class Admin {
  // Deletes an event
  private deleteEvent(eventItem: IEventItem, onRefresh: () => void) {
    // Clear the modal
    Modal.clear();

    // Set the header
    Modal.setHeader("Delete Event");

    // Set the body
    Modal.setBody("Are you sure you wanted to delete the selected event?");

    // Set the type
    Modal.setType(Components.ModalTypes.Medium);

    // Set the footer
    Modal.setFooter(Components.ButtonGroup({
      buttons: [
        {
          text: "Yes",
          type: Components.ButtonTypes.Primary,
          onClick: () => {
            // Show the loading dialog
            LoadingDialog.setHeader("Delete Event");
            LoadingDialog.setBody("Deleting the event");
            LoadingDialog.show();

            // Delete the item
            eventItem.delete().execute(
              () => {
                // Refresh the dashboard
                onRefresh();

                // Hide the dialog/modal
                Modal.hide();
                LoadingDialog.hide();
              },
              () => {
                LoadingDialog.hide();
              }
            );
          },
        },
        {
          text: "No",
          type: Components.ButtonTypes.Secondary,
          onClick: () => {
            // Hide the modal
            Modal.hide();
          },
        },
      ],
    }).el);

    // Show the modal
    Modal.show();
  }

  // Generates the navigation items
  generateNavItems(canEditEvent: boolean, onRefresh: () => void): Components.INavbarItem[] {
    let navItems: Components.INavbarItem[] = [];

    // See if this is the admin
    if (DataSource.IsAdmin) {
      // Add the new event option
      navItems.push({
        className: "btn-primary",
        isDisabled: !canEditEvent,
        text: " NEW EVENT",
        isButton: true,
        onClick: () => {
          // Create an item
          ItemForm.create({
            onUpdate: () => { onRefresh(); },
            onCreateEditForm: props => { return this.updateProps(props); },
            onFormButtonsRendering: buttons => { return this.updateFooter(buttons); }
          });

          // Update the modal
          Modal.setScrollable(true);
        },
        iconType: calendarPlus,
        iconSize: 18,
      });

      // Add the manage groups option
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
              // Show the manager's group
              window.open(DataSource.GetManagersUrl(), "_blank");
            },
          },
          {
            text: "Members",
            href: DataSource.GetManagersUrl(),
            onClick: () => {
              // Show the member's group
              window.open(DataSource.GetMembersUrl(), "_blank");
            },
          },
        ],
      });

      // Add an option to manage the application
      navItems.push({
        className: "btn-primary",
        text: "MANAGE APP",
        isButton: true,
        onClick: () => {
          // Show the installation dialog
          InstallationRequired.showDialog();
        }
      });
    }

    // Return the nav items
    return navItems;
  }

  // Edits the event
  private editEvent(eventItem: IEventItem, onRefresh: () => void) {
    // Display the edit form
    ItemForm.edit({
      itemId: eventItem.Id,
      onUpdate: () => { onRefresh(); },
      onCreateEditForm: props => { return this.updateProps(props); },
      onFormButtonsRendering: buttons => { return this.updateFooter(buttons); }
    });

    // Update the modal
    Modal.setScrollable(true);
  }

  // Renders the event menu
  renderEventMenu(el: HTMLElement, eventItem: IEventItem, canEditEvent: boolean, canDeleteEvent: boolean, onRefresh: () => void) {
    // Add the admin dropdown
    let adminDropdown = Components.Dropdown({
      el: el,
      className: "eventRegAdmin",
      items: [
        {
          isDisabled: !canEditEvent,
          text: " Edit",
          onClick: (button) => {
            this.editEvent(eventItem, onRefresh);
          },
        },
        {
          isDisabled: !canDeleteEvent,
          text: " Delete",
          onClick: (button) => {
            this.deleteEvent(eventItem, onRefresh);
          },
        },
        {
          text: " View Roster",
          onClick: (button) => {
            this.viewRoster(eventItem);
          },
        },
      ],
    });

    let adminDropdownEl = adminDropdown.el.querySelector("button");
    if (adminDropdownEl) {
      // Update the class
      adminDropdownEl.classList.add("btn-icon");
      adminDropdownEl.classList.add("w-100");

      // Append the icon
      adminDropdownEl.appendChild(gearWideConnected(16));
    }
  }

  // Updates the footer for the new/edit form
  private updateFooter(buttons: Components.IButtonProps[]) {
    // Update the default button
    buttons[0].type = Components.ButtonTypes.Primary;

    // Add the cancel button
    buttons.push({
      text: "Cancel",
      type: Components.ButtonTypes.Secondary,
      onClick: () => {
        ItemForm.close();
      }
    });

    // Return the buttons
    return buttons;
  }

  // Updates the new/edit form properties
  private updateProps(props: Components.IListFormEditProps) {
    // Set the rendering event
    props.onControlRendering = (ctrl, field) => {
      // See if this is the start date
      if (field.InternalName == "StartDate") {
        // Add validation
        ctrl.onValidate = (ctrl, results) => {
          // See if the start date is after the end date
          let startDate = results.value;
          let endDate = ItemForm.EditForm.getControl("EndDate").getValue();
          if (moment(startDate).isAfter(moment(endDate))) {
            // Update the validation
            results.isValid = false;
            results.invalidMessage = "Start Date cannot be after the End Date";
          }

          // Return the results
          return results;
        }
      }
    }

    // Return the properties
    return props;
  }

  // Displays the roster in a modal
  private viewRoster(eventItem: IEventItem) {
    //insert the print style css
    let modalCSS = [
      "@media print {",
      "\tbody * { visibility: hidden; }",
      "\t#core-modal, #core-modal * { visibility: visible; }",
      "\t#core-modal { position: absolute; left: 0; top: 0; }",
      "\t#core-modal .modal-footer, #core-modal .modal-footer *, #core-modal .modal-header, #core-modal .modal-header * { display: none; }",
      "}"
    ].join('\n');

    // Create the style element, if it does not exist
    if (document.getElementById("eventRegPrintStyle") == null) {
      let elStyle = document.createElement("style");
      elStyle.id = "eventRegPrintStyle";
      elStyle.innerHTML = modalCSS;
      document.head.appendChild(elStyle);
    }

    // Get the POCs
    let pocs = eventItem["POCId"] ? eventItem["POCId"].results : null || [];
    let pocString = "";
    for (let i = 0; i < pocs.length; i++) {
      let poc = Web().getUserById(pocs[i]).executeAndWait().Title;
      if (i > 0) pocString += "<br/>";
      pocString += poc;
    }

    // Get the list of registered usernames
    let usersRegistered = eventItem.RegisteredUsersId ? eventItem.RegisteredUsersId.results : null || [];
    let usersTable = `<tr><td colspan='4'>No registrations were found for the event</td></tr>`;
    for (let x = 0; x < usersRegistered.length; x++) {
      let userTitle = Web().getUserById(usersRegistered[x]).executeAndWait().Title;
      if (x == 0) usersTable = "";
      usersTable += `<tr><td>${x + 1}</td><td>${userTitle}</td><td></td><td></td></tr>`;
    }

    // Create the Jumbotron text
    let jumbotronContent = `
      <div class='table-responsive' id='print-view'>
        <table class='table table-striped'>
          <thead class='thead-dark'>
            <tr><th colspan='2' style='text-align:center'><h5>${eventItem.Title}</h5></th></tr>
          </thead>
          <tbody>
            <tr><td><strong>Start Date</strong></td><td> ${moment(eventItem.StartDate).format("MM-DD-YYYY HH:mm")}</td></tr>
            <tr><td><strong>End Date</strong></td><td>${moment(eventItem.EndDate).format("MM-DD-YYYY HH:mm")}</td></tr>
            <tr><td><strong>Location</strong></td><td>${eventItem.Location}</td></tr>
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

    // Clear the modal
    Modal.clear();

    // Set the type
    Modal.setType(Components.ModalTypes.Large);

    // Set the header
    Modal.setHeader(eventItem.Title + " - Event Registrations");

    // Set the body
    Modal.setBody(Components.Jumbotron({
      title: eventItem.Title,
      content: jumbotronContent,
    }).el);

    // Set the footer
    Modal.setFooter(Components.ButtonGroup({
      buttons: [
        {
          text: "Print",
          type: Components.ButtonTypes.Primary,
          onClick: () => {
            // Show the print dialog
            window.print();
          },
        },
        {
          text: "Close",
          type: Components.ButtonTypes.Secondary,
          onClick: () => {
            // Close the modal
            Modal.hide();
          },
        },
      ]
    }).el);

    // Show the modal
    Modal.show();
  }
}