import * as React from "react";
import Progress from "./Progress";

import ExcelApp from "./Hosts/ExcelApp";


export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
  host: Office.HostType;
  // account: any;
  // emailMessages: any;
  // error: string;
  // graphProfile: any;
  // onSignIn: () => void
  // onSignOut: () => void
  // onRequestEmailToken: () => void
}

export interface AppState {

}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {

    };
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div>
        {
          ((host: Office.HostType) => {
            switch (host) {
              case Office.HostType.Excel:
                return <ExcelApp></ExcelApp>;
              default:
                return <div>Wrong host</div>;
            }
          })(this.props.host)
        }
      </div>

    );
  }

}
