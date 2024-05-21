import { useContext, useState } from "react";
import {
  Image,
  TabList,
  Tab,
  SelectTabEvent,
  SelectTabData,
  TabValue,
} from "@fluentui/react-components";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { app, authentication, chat, Context, dialog, FrameContexts, HostClientType, geoLocation, getContext, location, pages, SdkError, tasks, people, version, webStorage } from "@microsoft/teams-js";
import { AzureFunctions } from "./AzureFunctions";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useData } from "@microsoft/teamsfx-react";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { TeamsFxContext } from "../Context";

function callInitialize() {
  app.initialize();
}

function onShareDeepLinkbutton() {
  pages.shareDeepLink({subPageId: "subPageId", subPageLabel: "subPageLabel"});
}

function onGetLocation() {
  location.getLocation({
    allowChooseLocation: true,
    showMap: false,
  },
  (error: SdkError, location: location.Location) => {
    console.log(`Location error: ${JSON.stringify(error)}`);
    console.log(`Location: ${JSON.stringify(location)}`);
  }
  );
}

async function onGetAuthToken() {
  try {
    const theToken = await authentication.getAuthToken();
    console.log(`Got the token`);
    console.log(theToken);
  } catch (error) {
    console.log(`Error getting token: ${error}`)
  }
}

function onLinkToSecondTab() {
  pages.navigateToApp({
    appId: window.location.hostname === "localhost" ? "3037e1e0-5b60-4350-bc2c-09ff2e4a17c7" : "1abc4bc4-c7c4-4f84-8ece-fc4a97d48149",
    pageId: "index1",
  });
}

function clearSubmissionAcknowledgement() {
  const dialogResultElement = document.getElementById("submissionAcknowledgement")!;
  dialogResultElement.innerText = "";
}

function openUrlDialog() {
  clearSubmissionAcknowledgement();
  dialog.url.open({
    url: window.location.href,
    title: "M365 Playground Dialog",
    size: { height: 600, width: 600 },
    },
    (result: dialog.ISdkResponse) => {
      const dialogResultElement = document.getElementById("submissionAcknowledgement")!;
      dialogResultElement.innerText = `Url Dialog submission occurred, result = ${result.result} err = ${result.err}`;
    }
  );
}

function writeToLocalStorage() {
  localStorage.setItem("myKey", "myValue");
}

function readFromLocalStorage() {
  const value = localStorage.getItem("myKey");
  const result = `Value read from local storage: ${value}`;
  console.log(result);
}

function selectPeople() {
  people.selectPeople({ singleSelect: true }).then((people: people.PeoplePickerResult[]) => {
    console.log(`People picker Success`);
  }).catch((error: SdkError) => {
    console.log(`People picker Error: ${error.errorCode}, message: ${error.message}`);
  });
}

function startSingleUserChat() {
  chat.openChat({user: "trharris@microsoft.com"});
}

function startGroupChat() {
  chat.openGroupChat({ users: ["trharris@microsoft.com", "erinha@microsoft.com"]});
}


function startAuthenticate() {
  authentication.authenticate({ 
    url: window.location.href,
    isExternal: false,
  })
}

const adaptiveCardJson = {
  "type": "AdaptiveCard",
  "body": [
      {
          "type": "TextBlock",
          "text": "Here is a ninja cat:"
      },
      {
          "type": "Image",
          "url": "http://adaptivecards.io/content/cats/1.png",
          "size": "Medium"
      }
  ],
  "actions": [
    {
        "type": "Action.Submit",
        "title": "Submit",
        "data": "Everything is awesome"
    }
  ],
  "version": "1.0"
};

const card2 = {
  type: "AdaptiveCard",
  body: [
   {
    type: "TextBlock",
    size: "Medium",
    weight: "Bolder",
    text: "Select user(s) in your organization."
   },
   {
    label: "1) Select user(s): ",
    isRequired: true,
    placeholder: "Search and select user(s)",
    type: "Input.ChoiceSet",
    choices: [],
    "choices.data": {
      type: "Data.Query",
      dataset: "graph.microsoft.com/users"
    },
    id: "selection",
    isMultiSelect: true,
    errorMessage: "Atleast one user must be selected."
   },
   {
    isRequired: true,
    label: "2) Message",
    type: "Input.Text",
    size:"Medium",
    placeholder: "Enter your message",
    id: 'message',
    errorMessage: "A message is required."
   },
   {
    type: "Input.Toggle",
    label: "3) Summary",
    title: "Include Summary?",
    valueOn: "1",
    valueOff: "2",
    value: "1",
    id: "sum_type"
   }
  ],
  actions: [
    {
      type: "Action.Submit",
      title: "Cancel",
      associatedInputs: "none"
    },
   {
    type: "Action.Submit",
    title: "Send"
   },
  ],
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  version: "1.4"
 };

function openCardAsObjectDialogAsTask() {
  clearSubmissionAcknowledgement();
  tasks.startTask({
    card: adaptiveCardJson as any as string,
  },
  (err: string, result: string | object) => {
    const dialogResultElement = document.getElementById("submissionAcknowledgement")!;
    dialogResultElement.innerText = `Card Dialog submission occurred, result = ${result} err = ${err}`;
  }
  )
}

function openCardDialogAsTask() {
  clearSubmissionAcknowledgement();
  tasks.startTask({
    card: JSON.stringify(adaptiveCardJson),    
  },
  (err: string, result: string | object) => {
    const dialogResultElement = document.getElementById("submissionAcknowledgement")!;
    dialogResultElement.innerText = `Card Dialog submission occurred, result = ${result} err = ${err}`;
  }
  )
}

function openCardDialogAsDialog() {
  clearSubmissionAcknowledgement();
  const dialogResultElement = document.getElementById("submissionAcknowledgement")!;

  try {
    dialog.adaptiveCard.open({
      card: JSON.stringify(adaptiveCardJson),
      size: { height: 600, width: 600 },
    },
    (result: dialog.ISdkResponse) => {
      dialogResultElement.innerText = `Card Dialog submission occurred, result = ${result.result} err = ${result.err}`;
    }
    );
  } catch (err) {
    dialogResultElement.innerText = `Exception thrown when opening card dialog as dialog, err = ${JSON.stringify(err)}`;
  }
}

function submitUrlDialog() {
  dialog.url.submit("Everything is super cool");
}

function navigateToSecondPage() {
  const currentUrl = new URL(window.location.href);
  const pathSegments = currentUrl.pathname.split('/');
  pathSegments[pathSegments.length - 1] = 'second.html';
  const newUrl = `${currentUrl.protocol}//${currentUrl.host}${pathSegments.join('/')}${currentUrl.search}${currentUrl.hash}`;
  window.location.href = newUrl;
}

function getWindowParentString() {
  var parent = window.parent;
  if (parent !== undefined && parent !== null) {
      if (parent === window.self) {
          return `PARENT IS SELF`;
      } else {
        try {
          return `PARENT IS NOT SELF, parent href = ${parent.location.href}`;
        } catch (e) {
          return `PARENT IS NOT SELF, parent location is not accessible: ${e}`;
        }
      }
  } else {
      return `PARENT IS UNDEFINED/NULL`;
  }
}

function getWindowOpenerString() {
  try {
    var opener = window.opener;
    if (opener !== undefined && opener !== null) {
        return opener.location.href;
    } else {
        return `OPENER IS UNDEFINED/NULL`;
    }
  } catch (err) {
      return `Exception while trying to access opener: ${err}`;
  }
}

export function Welcome(props: { showFunction?: boolean; environment?: string }) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const [selectedValue, setSelectedValue] = useState<TabValue>("local");

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value);
  };
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  const userName = loading || error ? "" : data!.displayName;
  const initResult = useData(async () => {
    await app.initialize();
    const context = await app.getContext();
    const appFrameContext = app.getFrameContext();
    return { appFrameContext, context };
  })?.data;

  getContext((contextv1: Context) => {
    const legacyContextHostTypeElement = document.getElementById("legacyContextHostType")!;
    legacyContextHostTypeElement.innerText = `Legacy context host type: ${contextv1.hostClientType}`;
  })
  
  const hubName: string | undefined = initResult?.context?.app.host.name;
  const clientType: HostClientType | undefined = initResult?.context?.app.host.clientType;
  const pageId: string | undefined = initResult?.context?.page.id;
  const frameContext: FrameContexts | undefined = initResult?.context?.page.frameContext;
  const appFrameContext = initResult?.appFrameContext;
  const cardDialogsIsSupported: boolean | undefined = initResult?.context === undefined ? undefined : dialog.adaptiveCard.isSupported();
  const locationSupported: boolean | undefined = initResult?.context === undefined ? undefined : location.isSupported();
  const pagesTabsSupported: boolean | undefined = initResult?.context === undefined ? undefined : pages.tabs.isSupported();
  const geoLocationSupported: boolean | undefined = initResult?.context === undefined ? undefined : geoLocation.isSupported();
  const peopleSupported: boolean | undefined = initResult?.context === undefined ? undefined : people.isSupported();
  const isWebStorageClearedOnUserLogOut: boolean | undefined = initResult?.context === undefined ? undefined : webStorage.isWebStorageClearedOnUserLogOut();

  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">WELCOME 1{userName ? ", " + userName : ""}!</h1>
        {hubName && (
          <p className="center">Your app is running in {hubName} on {clientType}</p>
        )}
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>
        <p className="center">TeamsJS version: {version}</p>
        <p className="center"><div id="legacyContextHostType">Legacy context host type: Not Retrieved Yet</div></p>
        <p className="center"><div id="currentContextHostType">Current context host type: {initResult?.context?.app.host.clientType}</div></p>
        <p className="center">Card Dialogs is supported: {cardDialogsIsSupported ? "true" : "false"}</p>
        <p className="center">Location is supported: {locationSupported ? "true" : "false"}</p>
        <p className="center">Pages.tabs is supported: {pagesTabsSupported ? "true" : "false"}</p>
        <p className="center">Geolocation is supported: {geoLocationSupported ? "true" : "false"}</p>
        <p className="center">People is supported: {peopleSupported ? "true" : "false"}</p>
        <p className="center">Is Web Storage Cleared on Logout? {isWebStorageClearedOnUserLogOut ? "true" : "false"}</p>
        {pageId && (
          <p className="center">The page id is {pageId}</p>
        )}
        <p className="center">The context frame context is {frameContext}</p>
        <p className="center">The app frame context is {appFrameContext}</p>
        <p className="center">The current URL is {window.location.href}</p>
        <p className="center">Window.parent is {getWindowParentString()}</p>
        <p className="center">Window.opener is {getWindowOpenerString()}</p>
        <p id="submissionAcknowledgement" className="center"></p>
        { frameContext === FrameContexts.content && (
          <div>
            <button onClick={openUrlDialog}>Open URL Dialog</button>
            <button onClick={openCardDialogAsTask}>Open Card (as string) Dialog (as task)</button>
            <button onClick={openCardAsObjectDialogAsTask}>Open Card (as object) Dialog (as task)</button>
            <button onClick={openCardDialogAsDialog}>Open Card Dialog (as dialog)</button>
          </div>
        )}
        { frameContext === FrameContexts.task && (
          <div>
            <button onClick={submitUrlDialog}>Submit Dialog</button>
          </div>
        )}
        <button onClick={callInitialize}>Initialize again</button>
        <button onClick={onGetAuthToken}>Get auth token</button>
        <button onClick={onShareDeepLinkbutton}>Share a deep link</button>
        <button onClick={onGetLocation}>Get Location</button>
        <button onClick={onLinkToSecondTab}>Link to Second Tab</button>
        <button onClick={writeToLocalStorage}>Write to Local Storage</button>
        <button onClick={readFromLocalStorage}>Read from Local Storage</button>
        <button onClick={selectPeople}>Select People</button>
        <button onClick={startSingleUserChat}>Start Single User Chat</button>
        <button onClick={startGroupChat}>Start Group User Chat</button>
        <button onClick={navigateToSecondPage}>Navigate to second page</button>
        <button onClick={() => window.location.href = "https://m365tab962ca2.z5.web.core.windows.net/index.html#/tab"}>Navigate to Cloud Deploy</button>
        <button onClick={() => window.location.href = "https://example2.com:53000/"}>Navigate to Example 2</button>
        <button onClick={startAuthenticate}>Authenticate</button>
        <button onClick={() => window.open("https://www.bing.com/")}>Open Bing in new window</button>
        <br></br>
        <a href="https://www.bing.com/" target="_blank" rel="noreferrer">Open Bing in new window</a>
        <br></br>
        <a href="https://www.example.com/">Open example.com in this window</a>
        <div className="tabList">
          <TabList selectedValue={selectedValue} onTabSelect={onTabSelect}>
            <Tab id="Local" value="local">
              1. Build your app locally
            </Tab>
            <Tab id="Azure" value="azure">
              2. Provision and Deploy to the Cloud
            </Tab>
            <Tab id="Publish" value="publish">
              3. Publish to Teams
            </Tab>
          </TabList>
          <div>
            {selectedValue === "local" && (
              <div>
                <EditCode showFunction={showFunction} />
                <CurrentUser userName={userName} />
                <Graph />
                {showFunction && <AzureFunctions />}
              </div>
            )}
            {selectedValue === "azure" && (
              <div>
                <Deploy />
              </div>
            )}
            {selectedValue === "publish" && (
              <div>
                <Publish />
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
