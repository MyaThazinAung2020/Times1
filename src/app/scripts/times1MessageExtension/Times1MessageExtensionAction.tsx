import * as React from "react";
import {
    PrimaryButton,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    Input,
    TeamsThemeContext,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the Times1MessageExtensionAction React component
 */
export interface ITimes1MessageExtensionActionState extends ITeamsBaseComponentState {
    email: string;
}

/**
 * Properties for the Times1MessageExtensionAction React component
 */
export interface ITimes1MessageExtensionActionProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Times1 Message Extension Task Module page
 */
export class Times1MessageExtensionAction extends TeamsBaseComponent<ITimes1MessageExtensionActionProps, ITimes1MessageExtensionActionState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };
        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Times1 Message Extension configuration</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
                                 <Input
                                    autoFocus
                                    placeholder="email@email.com"
                                    label="Enter an e-mail address"
                                    errorLabel={!this.state.email ? "This value is required" : undefined}
                                    value={this.state.email}
                                    onChange={(e) => {
                                        this.setState({
                                            email: e.target.value
                                        });
                                    }}
                                    required />
                            </div>
                            <div style={styles.section}>
                                <PrimaryButton onClick={() => {
                                    microsoftTeams.tasks.submitTask({
                                        email: this.state.email
                                    });
                                }}>OK</PrimaryButton>
                            </div>
                        </PanelBody>
                        <PanelFooter>
                            <div style={styles.footer}>
                                (C) Copyright MyaThazinAung
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
             </TeamsThemeContext.Provider>
        );
    }
}
