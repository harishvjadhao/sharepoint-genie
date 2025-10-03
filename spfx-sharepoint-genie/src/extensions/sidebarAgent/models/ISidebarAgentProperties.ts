export interface ISidebarAgentApplicationCustomizerProperties {}

export interface ISidebarAgentState {
  isPanelOpen: boolean;
  currentUserLogin: IUser;
  chatKey: number;
}

export interface ISidePanelProps {
  isOpen: boolean;
  properties: IWebpartSpecificProps;
  currentUserLogin: IUser;
  baseUrl: string;
  onDismiss: () => void;
  onNewConversation: () => void;
  chatKey: number;
}

export interface IWebpartSpecificProps {
  appClientId: string;
  tenantId: string;
  adScopes: string[];
  redirectUri: string;
  directConnectUrl: string;
  showTyping?: boolean;
  headerBackgroundColor?: string;
  agentTitle?: string;
}

export interface IUser {
  email: string;
  displayName: string;
  loginName: string;
}
