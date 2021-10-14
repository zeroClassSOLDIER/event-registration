import { Components, ContextInfo } from "gd-sprest-bs";
import { save2Fill } from "gd-sprest-bs/build/icons/svgs/save2Fill";
import * as moment from "moment";
import { IEventItem } from "./ds";

export class Calendar {
  private _item: IEventItem = null;
  private _isAdmin: boolean = false;
  private _el: HTMLElement = null;

  // Constructor
  constructor(el: HTMLElement, item: IEventItem, isAdmin: boolean) {
    // Save the properties
    this._el = el;
    this._item = item;
    this._isAdmin = isAdmin;

    // Render the calendar
    this.render();
  }

  // Generates a calendar url
  private getCalendarURL(item: IEventItem): string {
    // Create the ICS data
    let icsData = `BEGIN:VCALENDAR
CALSCALE:GREGORIAN
METHOD:PUBLISH
PRODID:-//Test Cal//EN
VERSION:2.0
BEGIN:VEVENT
UID:test-1
DTSTART;VALUE=DATE:${moment(item.StartDate).toDate().toISOString()}        
DTEND;VALUE=DATE:${moment(item.EndDate).toDate().toISOString()}        
SUMMARY:${item.Title}        
DESCRIPTION:${item.Description ?? ""}
LOCATION:${item.Location ?? ""}
END:VEVENT
END:VCALENDAR`;

    // Create the ICS file
    let data = new File([icsData], `${item.Title}_CourseRegistration`, {
      type: "text/plain",
    });

    // Release the data to prevent memory leaks and return the file
    return window.URL.createObjectURL(data);
  }

  // Renders the calendar
  private render() {
    let showAddToCalendar: boolean = false;

    // Determine if the user is currently registered
    let isRegistered = this._item.RegisteredUsersId ? this._item.RegisteredUsersId.results.indexOf(ContextInfo.userId) >= 0 : false;

    // See if this is an admin
    if (!this._isAdmin) {
      // See if they are registered
      if (isRegistered) showAddToCalendar = true;
    }

    // See if we are showing the calendar
    if (showAddToCalendar) {
      // Render the view tooltip
      let tooltip = Components.Tooltip({
        el: this._el,
        content: "Add event to calendar",
        placement: Components.TooltipPlacements.Top,
        btnProps: {
          iconType: save2Fill,
          iconSize: 24,
          toggle: "tooltip",
          type: Components.ButtonTypes.OutlinePrimary,
          href: showAddToCalendar ? this.getCalendarURL(this._item) : "",
        }
      });

      // Update the attribute
      tooltip.el.setAttribute("download", `${this._item.Title}.ics`);
    }
  }
}