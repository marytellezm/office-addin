import * as React from "react";
import { Spinner, SpinnerType } from "office-ui-fabric-react";
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import StartPageBody from "./StartPageBody";
import DocumentTitleEditor from "./DocumentTitleEditor"; // New component
import OfficeAddinMessageBar from "./OfficeAddinMessageBar";
import {
  logoutFromO365,
  signInO365,
} from "../../utilities/office-apis-helpers";
import { initializeAuth } from "../config/authConfig";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  authStatus?: string;
  headerMessage?: string;
  errorMessage?: string;
  isInitializing?: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      authStatus: "notLoggedIn",
      headerMessage: "Bienvenido",
      errorMessage: "",
      isInitializing: true, // New state to handle initialization
    };

    // Bind the methods that we want to pass to, and call in, a separate
    // module to this component. And rename setState to boundSetState
    // so code that passes boundSetState is more self-documenting.
    this.boundSetState = this.setState.bind(this);
    this.setToken = this.setToken.bind(this);
    this.setUserName = this.setUserName.bind(this);
    this.displayError = this.displayError.bind(this);
    this.login = this.login.bind(this);
  }

  /*
    Component Lifecycle
  */
  
  async componentDidMount() {
    // Check for existing authentication when component mounts
    if (this.props.isOfficeInitialized) {
      await this.checkExistingAuthentication();
    }
  }

  async componentDidUpdate(prevProps: AppProps) {
    // If Office just became initialized, check for existing auth
    if (!prevProps.isOfficeInitialized && this.props.isOfficeInitialized) {
      await this.checkExistingAuthentication();
    }
  }

  checkExistingAuthentication = async () => {
    try {
      this.setState({ isInitializing: true });
      
      // This will automatically set the auth state if user is already logged in
      await initializeAuth(
        this.boundSetState,
        this.setToken,
        this.setUserName,
        this.displayError
      );
      
    } catch (error) {
      console.error('Error during auth initialization:', error);
      this.setState({ 
        authStatus: "notLoggedIn",
        headerMessage: "Bienvenido" 
      });
    } finally {
      this.setState({ isInitializing: false });
    }
  };

  /*
        Properties
    */

  // The access token is not part of state because React is concerned with the
  // UI and the token is not used to affect the UI in any way.
  accessToken: string;
  userName: string;

  listItems: HeroListItem[] = [
    
    {
      icon: "Sync",
      primaryText: "De click en el botón Conectarse a Office 365 para iniciar sesión.",
    },
    {
      icon: "Save",
      primaryText: "Si el cartel de error no desaparece, haga clic en la X.",
    },
  ];

  /*
        Methods
    */

  boundSetState: () => {};

  setToken = (accesstoken: string) => {
    this.accessToken = accesstoken;
  };

  setUserName = (userName: string) => {
    this.userName = userName;
  };

  displayError = (error: string) => {
    this.setState({ errorMessage: error });
  };

  // Runs when the user clicks the X to close the message bar where
  // the error appears.
  errorDismissed = () => {
    this.setState({ errorMessage: "" });

    // If the error occurred during an "in process" phase (logging in),
    // the action didn't complete, so return the UI to the preceding state/view.
    this.setState((prevState) => {
      if (prevState.authStatus === "loginInProcess") {
        return { authStatus: "notLoggedIn" };
      }
      return null;
    });
  };

  login = async () => {
    await signInO365(
      this.boundSetState,
      this.setToken,
      this.setUserName,
      this.displayError
    );
  };

  logout = async () => {
    await logoutFromO365(
      this.boundSetState,
      this.setUserName,
      this.userName,
      this.displayError
    );
    
    // Clear local properties
    this.accessToken = '';
    this.userName = '';
  };

  // Remove the old getFileNames method as we no longer need it
  // All OneDrive functionality has been replaced with document title editing

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo="assets/logo-filled.png"
          message="Please sideload your add-in to see app body."
        />
      );
    }

    // Show spinner while checking for existing authentication
    if (this.state.isInitializing) {
      return (
        <div className="ms-welcome">
          <Header
            logo="assets/logo-filled.png"
            title={this.props.title}
            message="Initializing..."
          />
          <Spinner
            className="spinner"
            type={SpinnerType.large}
            label="Checking authentication status..."
          />
        </div>
      );
    }

    // Set the body of the page based on where the user is in the workflow.
    let body;

    if (this.state.authStatus === "notLoggedIn") {
      // Only show login button if user is not logged in
      body = <StartPageBody login={this.login} listItems={this.listItems} />;
    } else if (this.state.authStatus === "loginInProcess") {
      body = (
        <Spinner
          className="spinner"
          type={SpinnerType.large}
          label="Please sign-in on the pop-up window."
        />
      );
    } else if (this.state.authStatus === "loggedIn") {
      // User is authenticated, show document title editor
      body = <DocumentTitleEditor logout={this.logout} accessToken={this.accessToken} />;
    }

    return (
      <div>
        {this.state.errorMessage ? (
          <OfficeAddinMessageBar
            onDismiss={this.errorDismissed}
            message={this.state.errorMessage + " "}
          />
        ) : null}

        <div className="ms-welcome">
          <Header
            logo="assets/logo-filled.png"
            title={this.props.title}
            message={this.state.headerMessage}
            showLogout={this.state.authStatus === "loggedIn"}
            onLogout={this.logout}
          />
          {body}
        </div>
      </div>
    );
  }
}