import { Button } from "azure-devops-ui/Button";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
import { IWorkItemChangedArgs, IWorkItemFieldChangedArgs, IWorkItemFormService, IWorkItemLoadedArgs, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import * as SDK from "azure-devops-extension-sdk";
import * as React from "react";
import { showRootComponent } from "../../Common";
import "azure-devops-ui/Core/override.css";
import axios from "axios";

interface WorkItemFormGroupComponentState {
    eventContent: string;
    name: string;
    errorMessage: string;
}

export class WorkItemFormGroupComponent extends React.Component<{}, WorkItemFormGroupComponentState> {
    constructor(props: {}) {
        super(props);
        this.state = {
            eventContent: "",
            name: "",
            errorMessage: ""
        };
    }

    public componentDidMount() {
        SDK.init().then(() => {
            SDK.register(SDK.getContributionId(), () => {
                return {
                    // Called when the active work item is modified
                    onFieldChanged: (args: IWorkItemFieldChangedArgs) => {
                        this.setState({
                            eventContent: `onFieldChanged - ${JSON.stringify(args)}`
                        });
                    },

                    // Called when a new work item is being loaded in the UI
                    onLoaded: (args: IWorkItemLoadedArgs) => {
                        this.setState({
                            eventContent: `onLoaded - ${JSON.stringify(args)}`
                        });
                    },

                    // Called when the active work item is being unloaded in the UI
                    onUnloaded: (args: IWorkItemChangedArgs) => {
                        this.setState({
                            eventContent: `onUnloaded - ${JSON.stringify(args)}`
                        });
                    },

                    // Called after the work item has been saved
                    onSaved: (args: IWorkItemChangedArgs) => {
                        this.setState({
                            eventContent: `onSaved - ${JSON.stringify(args)}`
                        });
                    },

                    // Called when the work item is reset to its unmodified state (undo)
                    onReset: (args: IWorkItemChangedArgs) => {
                        this.setState({
                            eventContent: `onReset - ${JSON.stringify(args)}`
                        });
                    },

                    // Called when the work item has been refreshed from the server
                    onRefreshed: (args: IWorkItemChangedArgs) => {
                        this.setState({
                            eventContent: `onRefreshed - ${JSON.stringify(args)}`
                        });
                    }
                }
            });
        });
    }

    public render(): JSX.Element {
        return (
            <div>
                EventContext: {this.state.eventContent}
                <TextField
                    value={this.state.name}
                    onChange={(e, newValue) => (this.setState({ name: newValue }))}
                    placeholder="Enter your name"
                    width={TextFieldWidth.standard} />
                <Button
                    text="Ask the backend for a greeting"
                    primary={true}
                    onClick={() => this.onClick()}
                />
                {this.state.errorMessage}
            </div>
        );
    }

    private async onClick() {
        const url = "https://prod-93.eastus.logic.azure.com:443/workflows/16cad36640104aa7b67f43ca6f8cefba/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=k40FKhSc-0gkC7iulyBE6C8p3oF8A-0IaqyD86dsX0s";
        const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
        const jsonData = JSON.stringify({ID: (await workItemFormService.getId()).toString(), State: 'Active', Description: 'Teste'});
        const customConfig = {
            headers:{
                'Content-Type': 'application/json'
            }
        };

        const response = await axios.post(url, jsonData, customConfig);
        if (response.status === 200 || response.status == 202) {
            const workItemFormService = await SDK.getService<IWorkItemFormService>(
                WorkItemTrackingServiceIds.WorkItemFormService
            );
            workItemFormService.setFieldValue(
                "System.Description",
                `"${(await workItemFormService.getFieldValue("System.Title")).toString()}" set by extension`
            );
            this.setState({ errorMessage: "" });
        } else {
            this.setState({ errorMessage: JSON.stringify(response) });
        }

        

    }
}

export default WorkItemFormGroupComponent;

showRootComponent(<WorkItemFormGroupComponent />);