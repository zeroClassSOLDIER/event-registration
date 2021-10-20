import { ContextInfo } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Updates the strings for SPFx
export const setContext = (context) => {
    // Set the page context
    ContextInfo.setPageContext(context);

    // Update the values
    Strings.EventRegConfig = ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/eventreg-config.json";
    Strings.SourceUrl = ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/";
    Strings.SolutionUrl = ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/index.html";
}

// Strings
const Strings = {
    AppElementId: "event-registration",
    GlobalVariable: "EventRegistration",
    Lists: {
        Events: "Events"
    },
    ProjectName: "Event Registration",
    ProjectDescription: "Allows users to sign up for events.",
    EventRegConfig: ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/eventreg-config.json",
    SourceUrl: ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/",
    SolutionUrl: ContextInfo.webServerRelativeUrl + "/sites/dev/SiteAssets/Event-Registration/index.html",
    Version: "0.1",
};
export default Strings;