import * as React from "react";
import { Button, ButtonType, Announced, IPersonaProps, NormalPeoplePicker } from "office-ui-fabric-react";
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
  selectedUsers: IPersonaProps[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      error: "",
      userList: [],
      numberOfSuggestions: 0,
      selectedUsers: undefined
    };
  }

  async componentDidMount() {
    var me = this;
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
    Excel.run(async context => {
      // const sheet = context.workbook.worksheets.getItem("Users");
      // sheet.load("tables");
      // await context.sync();
      var selectedUsers: IPersonaProps[] = [];
      var table = context.workbook.tables.getItemOrNullObject("UsersTable");
      await context.sync();
      if (!table.isNullObject) {
        var tableRows = table.rows;

        await tableRows.load("items");
        await context.sync();
        if (tableRows) {
          for (var i = 0; i < tableRows.items.length; i++) {
            selectedUsers.push({
              imageUrl: `https://outlook.office.com/owa/service.svc/s/GetPersonaPhoto?email=${tableRows.items[i].values[0][2]}&UA=0&size=HR64x64&sc=1538493608488`,
              itemID: tableRows.items[i].values[0][1],
              text: tableRows.items[i].values[0][0]
            })
          }
        }
      }


      me.setState({ selectedUsers: selectedUsers });
    })

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
    // .then((val) => console.log(val))
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
          {this.state.selectedUsers != undefined && <NormalPeoplePicker
            onResolveSuggestions={this._onFilterChanged}
            onItemSelected={this._onSelectedUser}
            getTextFromItem={this._getTextFromItem}
            pickerSuggestionsProps={{
              suggestionsHeaderText: 'Suggested People',
              noResultsFoundText: 'No People Found', // this alert handles the case when there are no suggestions available,

            }}
            inputProps={{
              'aria-label': 'People Picker'
            }}
            onChange={this._onRemoveUser}
            defaultSelectedItems={this.state.selectedUsers}
          />}
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
  private _onSelectedUser = async (selected: IPersonaProps) => {

    this.updateTable(selected as User);
    return selected;
  }
  public _onRemoveUser = (selected: IPersonaProps[]) => {
    this._removeUserFromTable(selected);
    return selected;
  }

  private _removeUserFromTable = async (users: IPersonaProps[]) => {
    await Excel.run(async context => {
      var table = context.workbook.tables.getItemOrNullObject("UsersTable");
      await table.rows.load("items");
      await context.sync();
      // var deletedRow = table.rows.items.findIndex((val) => val.values[0][2] == user.key)
      var indexes: number[] = []
      table.rows.items.forEach((row, i) => {
        var index = users.findIndex((val) => row.values[0][1] == val.itemID)
        if (index == -1)
          indexes.push(i);

      });
      indexes.forEach(index => {
        table.rows.getItemAt(index).delete();
      });
      // .filter((val, i) => val.values[i][1] == user.key)
      await context.sync();

    })
  }

  private updateTable = async (user: User) => {
    await Excel.run(async context => {

      // const sheet = context.workbook.worksheets.getItem("Users");
      // sheet.load();
      var table = context.workbook.tables.getItemOrNullObject("UsersTable");
      await context.sync();
      if (table.isNullObject) {
        table = context.workbook.worksheets.getItem("Users").tables.add("A1:D1", true /*hasHeaders*/);
        table.name = "UsersTable";
        table.getHeaderRowRange().values = [["Display Name", "Id", "email", "Office Location"]];
      }

      table.rows.add(null, [[user.displayName, user.id, user.mail, user.officeLocation]]);

      await context.sync();
      // console.log(`The range address was ${range.address}.`);
    });
  }

  private _getTextFromItem = (item: IPersonaProps): string => {
    return item.text;
  };

  private _onFilterChanged = async (filterText: string, selectedTags: IPersonaProps[]) => {
    // if (filterText && this.state.emptyInput) {
    //   this.setState({ emptyInput: false });
    // } else if (!filterText && !this.state.emptyInput) {
    //   this.setState({ emptyInput: true });
    // }
    const options = new MSALAuthenticationProviderOptions(GRAPH_REQUESTS.LOGIN.scopes);
    const authProvider = new ImplicitMSALAuthenticationProvider(msalApp, options);

    const client = Client.initWithMiddleware(
      { authProvider });

    return client.api(`/users?$filter=startsWith(displayName,'${filterText}')`).get()
      .then((users: { value: User[] }) => {
        if (users.value.length > 0) {
          this.setState({ numberOfSuggestions: users.value.length });
        }

        return users.value.map<IPersonaProps>((u) => {
          return {
            ...u,
            imageUrl: `https://outlook.office.com/owa/service.svc/s/GetPersonaPhoto?email=${u.userPrincipalName}&UA=0&size=HR64x64&sc=1538493608488`,
            text: u.displayName,
            itemID: u.id
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
