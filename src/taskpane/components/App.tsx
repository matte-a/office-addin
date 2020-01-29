import * as React from "react";
import { Button, ButtonType, Announced, TagPicker, ITag } from "office-ui-fabric-react";
// import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
// import AuthProvider from "../../components/AuthProvider";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */
import { msalApp, GRAPH_REQUESTS } from '../../components/auth-utils'
import { Client } from "@microsoft/microsoft-graph-client";

import { ImplicitMSALAuthenticationProvider } from "../../../node_modules/@microsoft/microsoft-graph-client/lib/src/ImplicitMSALAuthenticationProvider";
import { MSALAuthenticationProviderOptions } from "../../../node_modules/@microsoft/microsoft-graph-client/lib/src/MSALAuthenticationProviderOptions";
import { User } from "@microsoft/microsoft-graph-types";


export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
  // account: any;
  // emailMessages: any;
  // error: string;
  // graphProfile: any;
  // onSignIn: () => void
  // onSignOut: () => void
  // onRequestEmailToken: () => void
}

export interface AppState {
  listItems: HeroListItem[];
  error: string;
  userList: any[];
  numberOfSuggestions: number;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      error: "",
      userList: [],
      numberOfSuggestions: 0
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
    msalApp.acquireTokenSilent(GRAPH_REQUESTS.LOGIN)
      .catch((err) => {
        switch (err.errorCode) {
          case "user_login_error":
            msalApp
              .loginPopup(GRAPH_REQUESTS.LOGIN)
              .catch(error => {
                this.setState({
                  error: error.message
                });
              });
            break;
          case "consent_required":
          case "interaction_required":
          case "login_required":
            msalApp.acquireTokenPopup(GRAPH_REQUESTS.LOGIN)
              .then((value) => {
                console.log(value);
              })
              .catch((err) => {
                console.log(err);
              })
            break;
        }
        console.log(err);

      }
      )
      .then((val) => console.log(val))
  }

  click = async () => {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      // <AuthProvider>
      <div className="ms-welcome">
        {/* <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" /> */}
        <HeroList message="Discover what Office Add-ins can do for you today!" items={[]}>
          <TagPicker
            onResolveSuggestions={this._onFilterChanged}
            onItemSelected={this._onSelectedUser}
            getTextFromItem={this._getTextFromItem}
            pickerSuggestionsProps={{
              suggestionsHeaderText: 'Suggested Tags',
              noResultsFoundText: 'No Color Tags Found' // this alert handles the case when there are no suggestions available
            }}
            inputProps={{
              'aria-label': 'Tag Picker'
            }}
          />
          {this.state.numberOfSuggestions > 0 &&
            < Announced message={`${this.state.numberOfSuggestions} result`}></Announced>
          }
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >  Run
          </Button>

        </HeroList>
      </div >
      // </AuthProvider>

    );
  }
  private _onSelectedUser = async (selected: ITag) => {

    this.updateExcel(selected as User);
    return selected;
  }

  private updateExcel = async (user: User) => {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "UsersTable";
      expensesTable.getHeaderRowRange().values = [["Display Name", "Id", "email", "Office Location"]];
      expensesTable.rows.add(null, [[user.displayName, user.id, user.mail, user.officeLocation]]);

      await context.sync();
      // console.log(`The range address was ${range.address}.`);
    });
  }

  private _getTextFromItem = (item: ITag): string => {
    return item.name;
  };

  private _onFilterChanged = async (filterText: string, selectedTags: ITag[]) => {
    // if (filterText && this.state.emptyInput) {
    //   this.setState({ emptyInput: false });
    // } else if (!filterText && !this.state.emptyInput) {
    //   this.setState({ emptyInput: true });
    // }
    selectedTags[0]
    const options = new MSALAuthenticationProviderOptions(GRAPH_REQUESTS.LOGIN.scopes);
    const authProvider = new ImplicitMSALAuthenticationProvider(msalApp, options);

    const client = Client.initWithMiddleware(
      { authProvider });

    return client.api(`/users?$filter=startsWith(displayName,'${filterText}')`).get()
      .then((users: { value: User[] }) => {
        if (users.value.length > 0) {
          this.setState({ numberOfSuggestions: users.value.length });
        }

        return users.value.map((u) => {
          return {
            ...u,
            name: u.displayName,
            key: u.id
          }
        })
      });
    // const filteredTags = filterText
    //   ? _testTags
    //     .filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0)
    //     .filter(tag => !this._listContainsDocument(tag, tagList))
    //   : [];

    // if (filteredTags.length !== this.state.numberOfSuggestions) {
    //   this.setState({ numberOfSuggestions: filteredTags.length });
    // }

    // return filteredTags;
  };



}
