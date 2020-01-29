import * as React from "react";
import {
    msalApp,
    requiresInteraction,
    fetchMsGraph,
    isIE,
    GRAPH_ENDPOINTS,
    GRAPH_SCOPES,
    GRAPH_REQUESTS
} from "./auth-utils";

// If you support IE, our recommendation is that you sign-in using Redirect APIs
const useRedirectFlow = isIE();
// const useRedirectFlow = true;
export interface AuthProviderState {
    account: any;
    error: any;
    emailMessages: any;
    graphProfile: any;
}
export interface AuthProviderProps {
}
export default class AuthProvider extends React.Component<any, AuthProviderState> {
    constructor(props) {
        super(props);

        this.state = {
            account: null,
            error: null,
            emailMessages: null,
            graphProfile: null
        };
    }

    acquireToken(request) {
        return msalApp.acquireTokenSilent(request).catch((error) => {
            // Call acquireTokenPopup (popup window) in case of acquireTokenSilent failure
            // due to consent or interaction required ONLY
            if (requiresInteraction(error.errorCode)) {
                return msalApp.acquireTokenPopup(request);
            }

            console.error('Non-interactive error:', error.errorCode)
            throw new Error('Non-interactive error:' + error.errorCode);
        });
    }

    async onSignIn(redirect) {
        if (redirect) {
            return msalApp.loginRedirect(GRAPH_REQUESTS.LOGIN);
        }

        const loginResponse = await msalApp
            .loginPopup(GRAPH_REQUESTS.LOGIN)
            .catch(error => {
                this.setState({
                    error: error.message
                });
            });

        if (loginResponse) {
            this.setState({
                account: loginResponse.account,
                error: null
            });

            const tokenResponse = await this.acquireToken(
                GRAPH_REQUESTS.LOGIN
            ).catch(error => {
                this.setState({
                    error: error.message
                });
            });

            if (tokenResponse) {
                const graphProfile = await fetchMsGraph(
                    GRAPH_ENDPOINTS.ME,
                    tokenResponse.accessToken
                ).catch(() => {
                    this.setState({
                        error: "Unable to fetch Graph profile."
                    });
                });

                if (graphProfile) {
                    this.setState({
                        graphProfile
                    });
                }

                if (tokenResponse.scopes.indexOf(GRAPH_SCOPES.MAIL_READ) > 0) {
                    return this.readMail(tokenResponse.accessToken);
                }
            }
        }
    }

    onSignOut() {
        msalApp.logout();
    }

    async onRequestEmailToken() {
        const tokenResponse = await this.acquireToken(
            GRAPH_REQUESTS.EMAIL
        ).catch(() => {
            this.setState({
                error: "Unable to acquire access token for reading email."
            });
        });

        if (tokenResponse) {
            return this.readMail(tokenResponse.accessToken);
        }
    }

    async readMail(accessToken) {
        const emailMessages = await fetchMsGraph(
            GRAPH_ENDPOINTS.MAIL,
            accessToken
        ).catch(() => {
            this.setState({
                error: "Unable to fetch email messages."
            });
        });

        if (emailMessages) {
            this.setState({
                emailMessages,
                error: null
            });
        }
    }

    async componentDidMount() {
        msalApp.handleRedirectCallback(error => {
            if (error) {
                const errorMessage = error.errorMessage ? error.errorMessage : "Unable to acquire access token.";
                // setState works as long as navigateToLoginRequestUrl: false
                this.setState({
                    error: errorMessage
                });
            }
        });

        const account = msalApp.getAccount();

        this.setState({
            account
        });

        if (account) {
            const tokenResponse = await this.acquireToken(
                GRAPH_REQUESTS.LOGIN
            );

            if (tokenResponse) {
                const graphProfile = await fetchMsGraph(
                    GRAPH_ENDPOINTS.ME,
                    tokenResponse.accessToken
                ).catch(() => {
                    this.setState({
                        error: "Unable to fetch Graph profile."
                    });
                });

                if (graphProfile) {
                    this.setState({
                        graphProfile
                    });
                }

                if (tokenResponse.scopes.indexOf(GRAPH_SCOPES.MAIL_READ) > 0) {
                    return this.readMail(tokenResponse.accessToken);
                }
            }
        }
    }

    render() {
        return (
            <AuthProvider
                {...this.props}
                account={this.state.account}
                emailMessages={this.state.emailMessages}
                error={this.state.error}
                graphProfile={this.state.graphProfile}
                onSignIn={() => this.onSignIn(useRedirectFlow)
                }
                onSignOut={() => this.onSignOut()}
                onRequestEmailToken={() => this.onRequestEmailToken()}
            />
        );
    }
};
