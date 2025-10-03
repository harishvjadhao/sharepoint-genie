import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
  ApplicationCustomizerContext,
} from "@microsoft/sp-application-base";
import * as React from "react";
import * as ReactDOM from "react-dom";
import * as strings from "SidebarAgentApplicationCustomizerStrings";
import SidePanel from "./components/SidePanel/SidePanel";
import {
  ISidebarAgentApplicationCustomizerProperties,
  ISidebarAgentState,
} from "./models/ISidebarAgentProperties";

const LOG_SOURCE: string = "SidebarAgentApplicationCustomizer";

class SidebarAgentComponent extends React.Component<
  {
    properties: ISidebarAgentApplicationCustomizerProperties;
    context: ApplicationCustomizerContext;
  },
  ISidebarAgentState
> {
  constructor(props: {
    properties: ISidebarAgentApplicationCustomizerProperties;
    context: ApplicationCustomizerContext;
  }) {
    super(props);
    const user = props.context?.pageContext?.user;
    this.state = {
      isPanelOpen: false,
      currentUserLogin: user ? user.loginName || user.email : undefined,
      chatKey: 0,
    };
  }

  private _togglePanel = (): void => {
    this.setState({ isPanelOpen: !this.state.isPanelOpen });
  };

  private _onPanelDismiss = (): void => {
    this.setState({ isPanelOpen: false });
  };

  private _startNewConversation = (): void => {
    this.setState((prevState) => ({
      chatKey: prevState.chatKey + 1,
    }));
  };

  private _getBaseUrl = (): string => {
    // Use SPFx context to get the base site URL
    const context = this.props.context;
    if (context?.pageContext?.web?.absoluteUrl) {
      return context.pageContext.web.absoluteUrl;
    }
    // Fallback to origin if context is not available
    return window.location.origin;
  };

  public render(): React.ReactElement {
    const { isPanelOpen } = this.state;
    const { properties } = this.props;

    const headerColor = properties.headerBackgroundColor || "white";
    const agentTitle =
      properties.agentTitle ||
      "SharePoint Genie" ||
      `${this.props.context.pageContext.web.title} Agent` ||
      "Copilot Studio Agent";
    properties.agentTitle = agentTitle; // Ensure agentTitle is set in properties for SidePanel

    const hasRequiredProps =
      properties.appClientId &&
      properties.tenantId &&
      (properties.directConnectUrl ||
        (properties.agentIdentifier && properties.environmentId));

    if (!hasRequiredProps) {
      return (
        <div
          style={{ padding: "10px", background: "#f3f2f1", color: "#d83b01" }}
        >
          <strong>Configuration Error:</strong> Missing required properties.
          Please configure appClientId, tenantId, and either directConnectUrl OR
          both agentIdentifier and environmentId.
        </div>
      );
    }

    return (
      <div>
        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            alignItems: "center",
            padding: "4px 12px",
            background: headerColor,
          }}
        >
          <button
            onClick={this._togglePanel}
            aria-label={
              isPanelOpen ? `Close ${agentTitle}` : `Open ${agentTitle}`
            }
            style={{
              background: "transparent",
              border: "none",
              borderRadius: "8px",
              width: "55px",
              height: "55px",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              transition: "all 0.2s ease",
              padding: "6px",
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor =
                "rgba(255, 255, 255, 0.2)";
              e.currentTarget.style.transform = "scale(1.05)";
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = "transparent";
              e.currentTarget.style.transform = "scale(1)";
            }}
          >
            <CopilotIcon />
          </button>
        </div>

        <SidePanel
          isOpen={isPanelOpen}
          properties={properties}
          currentUserLogin={this.state.currentUserLogin}
          baseUrl={this._getBaseUrl()}
          onDismiss={this._onPanelDismiss}
          onNewConversation={this._startNewConversation}
          chatKey={this.state.chatKey}
        />
      </div>
    );
  }
}

function CopilotIcon(): JSX.Element {
  return (
    <svg
      version="1.1"
      id="Layer_1"
      xmlns="http://www.w3.org/2000/svg"
      x="0"
      y="0"
      viewBox="0 0 512 512"
      xmlSpace="preserve"
    >
      <style>{`
        .st1{fill:#edf3fc}
        .st2{fill:#330d84}
        .st3{fill:#ffbe1b}
        .st10{fill:#5d8ef9}
      `}</style>
      <path
        className="st1"
        d="M255.999 40.928c-118.778 0-215.071 96.294-215.071 215.074 0 118.776 96.292 215.068 215.071 215.068S471.07 374.778 471.07 256.002c0-118.78-96.293-215.074-215.071-215.074z"
      />
      <path
        className="st1"
        d="M255.999 1C115.391 1 1 115.392 1 256.002 1 396.609 115.391 511 255.999 511S511 396.609 511 256.002C511 115.392 396.607 1 255.999 1zm0 501.832c-136.103 0-246.83-110.728-246.83-246.83 0-136.104 110.727-246.833 246.83-246.833 136.102 0 246.832 110.729 246.832 246.833 0 136.102-110.73 246.83-246.832 246.83z"
      />
      <path
        className="st3"
        d="m178.665 119.318 2.386.598-2.386.598a39.02 39.02 0 0 0-28.369 28.374l-.601 2.383-.599-2.383a39.022 39.022 0 0 0-28.376-28.374l-2.377-.598 2.377-.598a39.029 39.029 0 0 0 28.376-28.369l.599-2.381.601 2.381a39.026 39.026 0 0 0 28.369 28.369z"
      />
      <path
        className="st10"
        d="m223.217 399.809 2.386.598-2.386.598a39.02 39.02 0 0 0-28.369 28.374l-.601 2.383-.599-2.383a39.022 39.022 0 0 0-28.376-28.374l-2.377-.598 2.377-.598a39.029 39.029 0 0 0 28.376-28.369l.599-2.381.601 2.381a39.03 39.03 0 0 0 28.369 28.369z"
      />
      <path
        className="st2"
        d="m440.518 307.009 2.033.51-2.033.508a33.25 33.25 0 0 0-24.172 24.176l-.511 2.029-.51-2.029a33.252 33.252 0 0 0-24.177-24.176l-2.025-.508 2.025-.51a33.25 33.25 0 0 0 24.177-24.173l.51-2.027.511 2.027a33.25 33.25 0 0 0 24.172 24.173z"
      />
      <path
        className="st10"
        d="M297.973 298.495H182.992c-12.883 0-23.323-10.442-23.323-23.327v-96.296h-22.057c-12.881 0-23.322 10.445-23.322 23.32v103.229c0 12.885 10.441 23.327 23.322 23.327l-14.3 56.635 80.911-56.635h114.981a23.22 23.22 0 0 0 14.622-5.158l-35.853-25.095z"
      />
      <path
        className="st2"
        d="m364.585 298.495 14.3 56.635-80.911-56.635H182.992c-12.883 0-23.323-10.442-23.323-23.327V171.939c0-12.875 10.441-23.32 23.323-23.32h181.593c12.881 0 23.322 10.445 23.322 23.32v103.229c0 12.886-10.441 23.327-23.322 23.327z"
      />
      <circle
        transform="rotate(-22.5 219.366 220.851)"
        className="st3"
        cx="219.376"
        cy="220.86"
        r="17.136"
      />
      <circle
        transform="rotate(-22.5 273.774 220.851)"
        className="st3"
        cx="273.787"
        cy="220.86"
        r="17.136"
      />
      <circle
        transform="rotate(-22.5 328.183 220.852)"
        className="st3"
        cx="328.198"
        cy="220.86"
        r="17.136"
      />
      <path
        className="st10"
        d="m267.223 365.981 1.323.333-1.323.331a21.712 21.712 0 0 0-15.793 15.794l-.332 1.323-.334-1.323a21.715 21.715 0 0 0-15.791-15.794l-1.324-.331 1.324-.333a21.715 21.715 0 0 0 15.791-15.789l.334-1.327.332 1.327a21.712 21.712 0 0 0 15.793 15.789z"
      />
    </svg>
  );
}

export default class SidebarAgentApplicationCustomizer extends BaseApplicationCustomizer<ISidebarAgentApplicationCustomizerProperties> {
  private _topPlaceholder?: PlaceholderContent;
  private _reactContainer?: HTMLDivElement;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (!this.properties.appClientId || !this.properties.tenantId) {
      Log.error(
        LOG_SOURCE,
        new Error("appClientId and tenantId are required properties.")
      );
      return Promise.reject(
        "Missing required properties: appClientId and tenantId"
      );
    }

    if (
      !this.properties.directConnectUrl &&
      (!this.properties.agentIdentifier || !this.properties.environmentId)
    ) {
      Log.error(
        LOG_SOURCE,
        new Error(
          "Either directConnectUrl OR both agentIdentifier and environmentId must be provided."
        )
      );
      return Promise.reject(
        "Missing required properties: Either provide directConnectUrl OR both agentIdentifier and environmentId"
      );
    }

    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceholders
    );
    this._renderPlaceholders();
    return Promise.resolve();
  }

  private _renderPlaceholders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (!this._topPlaceholder) {
        Log.warn(LOG_SOURCE, "Top placeholder not available.");
        return;
      }

      if (this._topPlaceholder.domElement) {
        this._topPlaceholder.domElement.innerHTML = "";
        this._reactContainer = document.createElement("div");
        this._topPlaceholder.domElement.appendChild(this._reactContainer);

        const componentElement: React.ReactElement = React.createElement(
          SidebarAgentComponent,
          { properties: this.properties, context: this.context }
        );

        ReactDOM.render(componentElement, this._reactContainer);
      }
    }
  }

  private _onDispose = (): void => {
    if (this._reactContainer) {
      try {
        ReactDOM.unmountComponentAtNode(this._reactContainer);
      } catch {
        /* noop */
      }
      this._reactContainer = undefined;
    }
    Log.info(LOG_SOURCE, "Disposed Top placeholder content.");
  };
}
