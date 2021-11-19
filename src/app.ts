import { Dashboard, ItemForm, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes } from "gd-sprest-bs";
import { calendarEvent } from "gd-sprest-bs/build/icons/svgs/calendarEvent";
import * as jQuery from "jquery";
import * as moment from "moment";
import { Admin } from "./admin";
import { Calendar } from "./calendar";
import { DocumentsView } from "./documents"
import { DataSource, IEventItem } from "./ds";
import { Member } from "./member";
import { Registration } from "./registration";
import Strings from "./strings";


/**
 * Main Application
 */
export class App {
  //global vars
  private _canDeleteEvent: boolean;
  private _canEditEvent: boolean;
  private _canViewEvent: boolean;
  private _dashboard: Dashboard = null;
  private _el: HTMLElement = null;
  private _isAdmin: boolean = false;

  // Constructor
  constructor(el: HTMLElement) {
    // Set the list name
    ItemForm.ListName = Strings.Lists.Events;

    // Set the global variables
    this._canDeleteEvent = Helper.hasPermissions(DataSource.EventRegPerms, [SPTypes.BasePermissionTypes.DeleteListItems]);
    this._canEditEvent = Helper.hasPermissions(DataSource.EventRegPerms, [SPTypes.BasePermissionTypes.EditListItems]);
    this._canViewEvent = Helper.hasPermissions(DataSource.EventRegPerms, [SPTypes.BasePermissionTypes.ViewListItems]);
    this._el = el;

    // Get admin status
    this._isAdmin = DataSource.IsAdmin;

    // Render the dashboard
    this.render();
  }

  // Refreshes the dashboard
  private refresh() {
    // Show a loading dialog
    LoadingDialog.setHeader("Refreshing the Data");
    LoadingDialog.setBody("This will close after the data is loaded.");

    // Load the events
    DataSource.init().then(() => {
      // Clear the element
      while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

      // Render the dashboard
      this.render();

      // Hide the dialog
      LoadingDialog.hide();
    });
  }

  // Renders the dashboard
  private render() {
    let admin = new Admin();
    let member = new Member();

    // Create the dashboard
    this._dashboard = new Dashboard({
      el: this._el,
      useModal: true,
      hideHeader: DataSource.Configuration.hideHeader,
      header: {
        title: DataSource.Configuration.headerTitle || Strings.ProjectName,
        onRendered: (el) => {
          // See if the image url is defined
          if (DataSource.Configuration.headerImage) {
            // Update the header
            el.style.backgroundImage = "url(" + DataSource.Configuration.headerImage + ")";
          }
        }
      },
      filters: {
        items: [
          {
            header: "Event Status",
            items: DataSource.StatusFilters,
            onFilter: (value: string) => {
              // Update the default flag
              DataSource.StatusFilters[0].isSelected = value ? true : false;

              // Filter the dashboard
              this._dashboard.filter(0, value ? "" : "Active");
            },
          },
        ],
      },
      navigation: {
        showFilter: this._isAdmin,
        items: admin.generateNavItems(this._canEditEvent, () => { this.refresh(); }),
      },
      footer: {
        itemsEnd: [
          {
            text: "v" + Strings.Version,
          },
        ],
      },
      table: {
        rows: DataSource.Events,
        dtProps: {
          dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
          columnDefs: [
            {
              targets: [0, 6, 10, 11, 12],
              orderable: false,
              searchable: false,
            },
            { width: "10%", targets: [3, 4] },
            {
              targets: 2, render: function (data, type, row) {
                // Limit the length of the Description column to 50 chars
                let esc = function (t) {
                  return t
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;');
                };
                // Order, search and type get the original data
                if (type !== 'display') { return data; }
                if (typeof data !== 'number' && typeof data !== 'string') { return data; }
                data = data.toString(); // cast numbers
                if (data.length < 50) { return data; }

                // Find the last white space character in the string
                let trunc = esc(data.substr(0, 50).replace(/\s([^\s]*)$/, ''));
                return '<span title="' + esc(data) + '">' + trunc + '&#8230;</span>';
              }
            },
            this._isAdmin || DataSource.Configuration.hideAddToCalendarColumn == true ? { targets: [12], visible: false } : null
          ],
          // Add some classes to the dataTable elements
          drawCallback: function () {
            jQuery(".table", this._table).removeClass("no-footer");
            jQuery(".table", this._table).addClass("tbl-footer");
            jQuery(".table", this._table).addClass("table-striped");
            jQuery(".table thead th", this._table).addClass("align-middle");
            jQuery(".table tbody td", this._table).addClass("align-middle");
            jQuery(".dataTables_info", this._table).addClass("text-center");
            jQuery(".dataTables_length", this._table).addClass("pt-2");
            jQuery(".dataTables_paginate", this._table).addClass("pt-03");
          },
          // Sort descending by Start Date
          order: [[3, "asc"]],
          language: {
            emptyTable: "No events were found",
          },
        },
        columns: [
          {
            // 0 - View Event tooltip
            name: "",
            title: "",
            onRenderCell: (el, column, item: IEventItem) => {
              // Set the filter/search value
              let today = moment();
              let startDate = item.StartDate;
              let isActive = moment(startDate).isAfter(today);
              el.setAttribute("data-search", isActive ? "Active" : "Past");

              // Render the tooltip
              Components.Tooltip({
                el: el,
                content: "View Event",
                placement: Components.TooltipPlacements.Top,
                btnProps: {
                  isDisabled: !this._canViewEvent,
                  iconType: calendarEvent,
                  iconSize: 24,
                  toggle: "tooltip",
                  type: Components.ButtonTypes.OutlinePrimary,
                  onClick: (button) => {
                    ItemForm.view({
                      itemId: item.Id,
                      useModal: true,
                      onCreateViewForm: (props) => {
                        props.excludeFields = [
                          "RegisteredUsers",
                          "WaitListedUsers",
                        ];
                        return props;
                      },
                      onSetFooter: (elFooter) => {
                        // Render the close button
                        Components.Button({
                          el: elFooter,
                          text: "Close",
                          type: Components.ButtonTypes.Secondary,
                          onClick: (button) => {
                            ItemForm.close();
                          }
                        });
                      }
                    });
                    document.querySelector(".modal-dialog").classList.add("modal-dialog-scrollable");
                  },
                }
              });
            },
          },
          {
            // 1 - Title
            name: "Title",
            title: "Title",
            onRenderCell: (el, column, item: IEventItem) => {
              // Determine the # of hours until the event starts
              let currDate = moment();
              let startDate = moment(item["StartDate"]);
              let resetHour = '12:00:00 am';
              let time = moment(resetHour, 'HH:mm:ss a');
              let currReset = currDate.set({ hour: time.get('hour'), minute: time.get('minute') });
              let startReset = startDate.set({ hour: time.get('hour'), minute: time.get('minute') });
              let dateDiff = moment(startReset, "DD/MM/YYYY").diff(moment(currReset, "DD/MM/YYYY"), "hours");
              //console.log("dateDiff: " + moment(currReset).format("DD MMM YYYY hh:mm a") + " - " + moment(startReset).format("DD MMM YYYY hh:mm a") + " = " + dateDiff);

              // See if this event is cancelled
              if (item.IsCancelled) {
                // Add a break after title
                let elBreak = document.createElement("br");
                el.appendChild(elBreak);

                // Render a badge
                Components.Badge({
                  el,
                  content: "CANCELLED",
                  isPill: true,
                  type: Components.BadgeTypes.Info
                });
              }
              // Else, see if the event starts w/in 24 hours
              else if (dateDiff <= 24 && dateDiff > 0) {
                // Update the style
                el.style.fontWeight = "bold";

                // Add a break after title
                let elBreak = document.createElement("br");
                el.appendChild(elBreak);

                // Render a badge
                Components.Badge({
                  el,
                  content: "LAST DAY TO REGISTER!",
                  isPill: true,
                  type: Components.BadgeTypes.Danger
                });
              }
            },
          },
          {
            // 2 - Description
            name: "Description",
            title: "Description",
          },
          {
            // 3 - Start Date
            name: "StartDate",
            title: "Start Date",
            onRenderCell: (el, column, item: IEventItem) => {
              let date = item[column.name];
              el.innerHTML =
                moment(date).format("MMMM DD, YYYY") +
                "<br>" +
                moment(date).format("dddd HH:mm");
              // Set the date/time filter/sort values
              el.setAttribute("data-filter", moment(item[column.name]).format("dddd MMMM DD YYYY"));
              el.setAttribute("data-sort", item[column.name]);
            },
          },
          {
            // 4 - End Date
            name: "EndDate",
            title: "End Date",
            onRenderCell: (el, column, item: IEventItem) => {
              let date = item[column.name];
              el.innerHTML =
                moment(date).format("MMMM DD, YYYY") +
                "<br/>" +
                moment(date).format("dddd HH:mm");
              // Set the date/time filter/sort values
              el.setAttribute("data-filter", moment(item[column.name]).format("dddd MMMM DD YYYY"));
              el.setAttribute("data-sort", item[column.name]);
            },
          },
          {
            // 5 - Location
            name: "Location",
            title: "Location",
          },
          {
            // 6 - Documents
            name: "",
            title: "Documents",
            onRenderCell: (el, column, item: IEventItem) => {
              // Render the document column
              new DocumentsView(el, item, this._isAdmin, this._canEditEvent, () => { this.refresh(); });
            },
          },
          {
            // 7- Open Spots
            name: "",
            title: "Open Spots",
            onRenderCell: (el, column, item: IEventItem) => {
              let capacity: number = item.Capacity
                ? (parseInt(item.Capacity) as number)
                : 0;
              let numUsers: number = item.RegisteredUsersId
                ? item.RegisteredUsersId.results.length
                : 0;
              let numWaitlisted: number = item.WaitListedUsersId
                ? item.WaitListedUsersId.results.length
                : 0;
              let eventFull: boolean = capacity == numUsers ? true : false;
              if (eventFull) {
                // Render the badges
                Components.Badge({
                  el: el,
                  type: Components.BadgeTypes.Danger,
                  content: "Full",
                  className: "w-50",
                });

                // Render a new line
                let breakLine = document.createElement("br");
                el.appendChild(breakLine);

                // Render a badge for the waitlisted users
                Components.Badge({
                  el: el,
                  type: Components.BadgeTypes.Primary,
                  content: numWaitlisted + " Waitlisted",
                  className: "w-100",
                });
              } else {
                // Render the capacity
                el.innerHTML = (capacity - numUsers).toString();

                // See if wait users exist
                if (numWaitlisted > 0) {
                  // Render a new line
                  let breakLine = document.createElement("br");
                  el.appendChild(breakLine);

                  // Render a badge for the waitlisted users
                  Components.Badge({
                    el: el,
                    type: Components.BadgeTypes.Primary,
                    content: numWaitlisted + " Waitlisted",
                    className: "w-100",
                  });
                }
              }
            },
          },
          {
            // 8 - Capacity
            name: "Capacity",
            title: "Capacity",
          },
          {
            // 9 - POC
            name: "",
            title: "POC",
            onRenderCell: (el, column, item: IEventItem) => {
              let pocs = ((item["POC"] ? item["POC"].results : null) || []).sort((a, b) => {
                if (a.Title < b.Title) { return -1; }
                if (a.Title > b.Title) { return 1; }
                return 0;
              });
              for (let i = 0; i < pocs.length; i++) {
                if (i > 0) el.innerHTML += "<br/>";
                el.innerHTML += pocs[i].Title;
              }
            },
          },
          {
            // 10 - Registration
            name: "",
            title: " Registration",
            onRenderCell: (el, column, item: IEventItem) => {
              if (item.IsCancelled) { return; }
              new Registration(el, item, () => { this.refresh(); });
            },
          },
          {
            // 11 - User/Admin Options
            name: "",
            title: this._isAdmin ? "Manage Event" : "Event Options",
            onRenderCell: (el, column, item: IEventItem) => {
              if (this._isAdmin) {
                // Render the admin menu
                admin.renderEventMenu(el, item, this._canEditEvent, this._canDeleteEvent, () => { this.refresh(); });
              } else {
                // Render the user menu
                member.renderEventMenu(el, item);
              }
            },
          },
          {
            // 12 - Add to Calendar
            name: "",
            title: "",
            onRenderCell: (el, column, item: IEventItem) => {
              new Calendar(el, item, this._isAdmin);
            }
          }
        ]
      }
    });

    // See if we are filtering active items
    if (this._dashboard.getFilter("Event Status").getValue() == null) {
      // Filter the dashboard
      this._dashboard.filter(0, "Active");
    }
  }
}