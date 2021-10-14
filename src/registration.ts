import { LoadingDialog } from "dattatable";
import { Components, ContextInfo, Utility, Web } from "gd-sprest-bs";
import { IEventItem } from "./ds";
import * as moment from "moment";
import { calendarPlusFill } from "gd-sprest-bs/build/icons/svgs/calendarPlusFill"
import { calendarMinusFill } from "gd-sprest-bs/build/icons/svgs/calendarMinusFill"
import { personPlusFill } from "gd-sprest-bs/build/icons/svgs/personPlusFill";
import { personXFill } from "gd-sprest-bs/build/icons/svgs/personXFill";

export class Registration {
    private _el: HTMLElement = null;
    private _item: IEventItem = null;
    private _onRefresh: () => void = null;

    constructor(el: HTMLElement, item: IEventItem, onRefresh: () => void) {
        this._el = el;
        this._item = item;
        this._onRefresh = onRefresh;
        this.render();
    }

    private render() {
        //see if the user if registered
        let isRegistered = this._item.RegisteredUsersId ? this._item.RegisteredUsersId.results.indexOf(ContextInfo.userId) >= 0 : false;
        //see if the course is full
        let numUsers: number = this._item.RegisteredUsersId ? this._item.RegisteredUsersId.results.length : 0;
        let capacity: number = this._item.Capacity ? (parseInt(this._item.Capacity) as number) : 0;
        let eventFull: boolean = numUsers == capacity ? true : false;
        //check if user is on the waitlist
        let userID = ContextInfo.userId;
        let userOnWaitList = this._item.WaitListedUsersId ? this._item.WaitListedUsersId.results.indexOf(ContextInfo.userId) >= 0 : false;
        // Render the buttons based on user/event status
        let btnText: string = "";
        let btnType: number = null;
        let dlg: string = "";
        let iconType: any = null;
        let iconSize: number = 24;
        let registerUserFromWaitlist: boolean = false;
        let userFromWaitlist: number = 0;

        if (userOnWaitList) {
            btnText = "Remove From Waitlist";
            btnType = Components.ButtonTypes.OutlineDanger;
            dlg = "Removing User From Waitlist";
            iconType = personXFill;
        } else if (eventFull && !isRegistered) {
            btnText = "Add To Waitlist";
            btnType = Components.ButtonTypes.OutlinePrimary;
            dlg = "Adding User To Waitlist";
            iconType = personPlusFill;
        } else if (isRegistered) {
            btnText = "Unregister";
            btnType = Components.ButtonTypes.OutlineDanger;
            dlg = "Unregistering the User";
            iconType = calendarMinusFill;
        } else {
            btnText = "Register";
            btnType = Components.ButtonTypes.OutlinePrimary;
            dlg = "Registering the User";
            iconType = calendarPlusFill;
        }

        let btn = Components.Tooltip({
            el: this._el,
            content: btnText,
            btnProps: {
                iconType: iconType,
                iconSize: iconSize,
                type: btnType,
                onClick: (button) => {
                    // See if the user is unregistered
                    let isUnregistered = btnText == "Unregister";
                    let userIsRegistering = btnText == "Register";

                    // Display a loading dialog
                    LoadingDialog.setHeader(dlg);
                    LoadingDialog.setBody(dlg);
                    LoadingDialog.show();

                    let waitListedUserIds = this._item.WaitListedUsersId ? this._item.WaitListedUsersId.results : [];
                    let registeredUserIds = this._item.RegisteredUsersId ? this._item.RegisteredUsersId.results : [];

                    // User is being removed from the waitlist
                    if (userOnWaitList) {
                        // Remove the user from the waitlist
                        let userIdx = waitListedUserIds.indexOf(userID);
                        waitListedUserIds.splice(userIdx, 1);
                    } else if (eventFull && !isRegistered) {
                        // Add the user to the waitlist
                        waitListedUserIds.push(ContextInfo.userId);
                    } else {
                        //User is registered or unregistered
                        // See if the user is unregistered
                        if (isUnregistered) {
                            // Get the user ids
                            let userIdx = registeredUserIds.indexOf(
                                ContextInfo.userId
                            );

                            // Remove the user
                            registeredUserIds.splice(userIdx, 1);

                            //if the event was full, add the next waitlist user
                            if (eventFull && waitListedUserIds.length > 0) {
                                userFromWaitlist = waitListedUserIds[0];
                                let idx = waitListedUserIds.indexOf(userFromWaitlist);
                                //remove from waitlist
                                waitListedUserIds.splice(idx, 1);
                                //add to registered users
                                registeredUserIds.push(userFromWaitlist);
                                registerUserFromWaitlist = true;
                            }
                        } else {
                            // Add the user
                            registeredUserIds.push(ContextInfo.userId);
                        }
                    }

                    // Update the list item
                    let updateFields = {};
                    if ((eventFull || userOnWaitList) && !isRegistered) {
                        updateFields = {
                            WaitListedUsersId: { results: waitListedUserIds },
                        };
                    } else if (isUnregistered) {
                        // Increase the open spots
                        let openSpots: number = parseInt(this._item.OpenSpots) as number;
                        if (registerUserFromWaitlist) {
                            updateFields = {
                                RegisteredUsersId: { results: registeredUserIds },
                                WaitListedUsersId: { results: waitListedUserIds },
                            };
                        } else {
                            updateFields = {
                                OpenSpots: openSpots - 1,
                                RegisteredUsersId: { results: registeredUserIds },
                            };
                        }
                    } //unregistering
                    else {
                        let openSpots: number = parseInt(this._item.OpenSpots) as number;
                        // decrease the open spots
                        updateFields = {
                            OpenSpots: openSpots + 1,
                            RegisteredUsersId: { results: registeredUserIds },
                        };
                    }
                    this._item.update(updateFields).execute(
                        // Success
                        () => {
                            //Send email
                            this.sendMail(userFromWaitlist, userIsRegistering);
                            // Hide the dialog
                            LoadingDialog.hide();

                            // Refresh the dashboard
                            this._onRefresh();
                        },
                        // Error
                        () => {
                            // TODO
                        }
                    );
                }
            }
        });
    }

    private sendMail(userFromWaitlist: number, userIsRegistering: boolean) {
        let userEmail = userFromWaitlist > 0 ? Web().getUserById(userFromWaitlist).executeAndWait().Email : ContextInfo.userEmail;
        let body = `${ContextInfo.userDisplayName}, you have successfully ${userFromWaitlist > 0 ? "been added from the waitlist" : "registered"} for the following event:
        <p><strong>Title:</strong>${this._item.Title}</p></br>
        <p><strong>Description:</strong>${this._item.Description}</p></br>
        <p><strong>Start Date:</strong>${moment(this._item.StartDate).format("MM-DD-YYYY HH:mm")}</p></br>
        <p><strong>End Date:</strong>${moment(this._item.EndDate).format("MM-DD-YYYY HH:mm")}</p></br>
        <p><strong>Location:</strong>${this._item.Location}`;
        let subject = `Successfully ${userFromWaitlist > 0 ? "added from the waitlist" : "registered"} for the event: ${this._item.Title}`;
        if (
            (userEmail && userIsRegistering) || userFromWaitlist > 0) {
            Utility()
                .sendEmail({
                    To: [userEmail],
                    Subject: subject,
                    Body: body,
                })
                .execute(
                    () => {
                        console.log("Successfully sent email");
                    },
                    () => {
                        console.error("Error sending email");
                    }
                );
        }
    }
}