import * as React from "react";
import styles from "./Errouiwpnpwp.module.scss";
import { IErrouiwpnpwpProps } from "./IErrouiwpnpwpProps";
import { IErrouiwpnpwpState } from "./IErrouiwpnpwpState";
import { escape } from "@microsoft/sp-lodash-subset";
import { PnpServices } from "../services/pnpservices";
//Step 1 : Imporeted all the required classess from office ui fabric
import {
  TextField,
  Dropdown,
  Selection,
  IDropdownStyles,
  IDropdownOption,
  IIconProps,
  PrimaryButton,
  Stack,
  IStackProps,
  IStackStyles,
  DetailsList,
  CheckboxVisibility,
  SelectionMode,
  DetailsListLayoutMode,
} from "office-ui-fabric-react";
import { LIST_COLUMNS } from "../shared/constants";
import { IListItem } from "../models/IListItem";
//step2

const stackTokens = { childrenGap: 50 };

const AddICon: IIconProps = { iconName: "Add" };
const ReadIcon: IIconProps = { iconName: "BulletedListText" };
const SaveIcon: IIconProps = { iconName: "Save" };
const DeleteIcon: IIconProps = { iconName: "Delete" };

const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

const ddlBatchOptions: IDropdownOption[] = [
  { key: "Batch 1", text: "Batch 1" },
  { key: "Batch 2", text: "Batch 2" },
  { key: "Batch 3", text: "Batch 3" },
];

const ddlLevelOFKnowledgeOptions: IDropdownOption[] = [
  { key: "Beginner", text: "Beginner" },
  { key: "Intermediate", text: "Intermediate" },
  { key: "Expert", text: "Expert" },
];

const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};

export default class Errouiwpnpwp extends React.Component<
  IErrouiwpnpwpProps,
  IErrouiwpnpwpState
> {
  private _sp: PnpServices;
  private _selection: Selection;
  constructor(props: IErrouiwpnpwpProps, state: IErrouiwpnpwpState) {
    super(props);

    this.state = {
      status: "Ready",
      ListItems: [],
      ListItem: {
        Id: 0,
        Title: "",
        Email: "",
        Batch: "",
        LevelOfKnowledge: "",
      },
    };

    this._sp = new PnpServices(this.props.context);
    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ ListItem: this._onItemSelectionChanged() }),
    });
  }

  private _onItemSelectionChanged(): any {
    const selectedItem = this._selection.getSelection()[0] as IListItem;

    return selectedItem;
  }

  private async callAndBindDetailsList(message: string): Promise<any> {
    await this._sp.getItems(this.props.listName).then((listItems) => {
      this.setState({
        ListItems: listItems,
        status: message,
      });
    });
  }

  private async _createItem(): Promise<any> {
    await this._sp
      .CreateItem(this.props.listName, this.state.ListItem)
      .then((Id) => {
        this.callAndBindDetailsList(
          "New Item Created Successfully with ID " + Id
        );
      });
  }

  private async _readItem(): Promise<any> {
    await this.callAndBindDetailsList("Items Loaded Successfully");
  }

  private async _updateItem(): Promise<any> {
    await this._sp
      .updateItem(this.props.listName, this.state.ListItem.Id, {
        Title: this.state.ListItem.Title,
        Email: this.state.ListItem.Email,
        Batch: this.state.ListItem.Batch,
        LevelOfKnowledge: this.state.ListItem.LevelOfKnowledge,
      })
      .then((Id) => {
        this.callAndBindDetailsList(`Item ${Id} Updated Successfully`);
      });
  }

  private async _deleteItem(): Promise<any> {
    try {
      await this._sp
        .deleteItem(this.props.listName, this.state.ListItem.Id)
        .then(() => {
          this.setState({ status: "Item Deleted Successfully" });
        });
    } catch (error) {}
  }

  componentDidMount(): void {
    this.callAndBindDetailsList("Record Loaded");
  }

  public render(): React.ReactElement<IErrouiwpnpwpProps> {
    return (
      <div>
        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <TextField
              label="Username"
              placeholder="Please enter username"
              value={this.state.ListItem.Title}
              onChange={(e, newValue) => {
                this.setState(
                  (state) => ((state.ListItem.Title = newValue), state)
                );
              }}
            />
            <TextField
              label="Email"
              placeholder="Please enter username"
              value={this.state.ListItem.Email}
              onChange={(e, newValue) => {
                this.setState(
                  (state) => ((state.ListItem.Email = newValue), state)
                );
              }}
            />

            <Dropdown
              placeholder="Select an option"
              label="Select Batch"
              options={ddlBatchOptions}
              styles={dropdownStyles}
              selectedKey={this.state.ListItem.Batch}
              defaultValue={this.state.ListItem.Batch}
              onChange={(e, newValue) => {
                this.setState(
                  (state) => ((state.ListItem.Batch = newValue.text), state)
                );
              }}
            />

            <Dropdown
              placeholder="Select an option"
              label="Select Level Of Knowledge"
              options={ddlLevelOFKnowledgeOptions}
              styles={dropdownStyles}
              selectedKey={this.state.ListItem.LevelOfKnowledge}
              defaultValue={this.state.ListItem.LevelOfKnowledge}
              onChange={(e, newValue) => {
                this.setState(
                  (state) => (
                    (state.ListItem.LevelOfKnowledge = newValue.text), state
                  )
                );
              }}
            />
          </Stack>
        </Stack>

        <hr />
        <Stack horizontal tokens={stackTokens}>
          <PrimaryButton
            text="Create"
            iconProps={AddICon}
            onClick={(e) => this._createItem()}
          />
          <PrimaryButton
            text="Read"
            iconProps={ReadIcon}
            onClick={(e) => this._readItem()}
          />
          <PrimaryButton
            text="Update"
            iconProps={SaveIcon}
            onClick={(e) => this._updateItem()}
          />
          <PrimaryButton
            text="Delete"
            iconProps={DeleteIcon}
            onClick={(e) => this._deleteItem()}
          />
        </Stack>

        <div id="divStatus">{this.state.status}</div>

        <hr />

        <DetailsList
          items={this.state.ListItems}
          columns={LIST_COLUMNS}
          setKey="Id"
          checkboxVisibility={CheckboxVisibility.onHover}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selection={this._selection}
        />
      </div>
    );
  }
}
