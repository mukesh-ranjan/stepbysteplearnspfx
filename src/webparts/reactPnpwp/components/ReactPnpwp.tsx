import * as React from "react";
import styles from "./ReactPnpwp.module.scss";
import { IReactPnpwpProps } from "./IReactPnpwpProps";
import { IReactPnpwpState } from "./IReactPnpwpState";
import { escape } from "@microsoft/sp-lodash-subset";
import { IListItem } from "../models/IListItem";
import { LIST_COLUMNS, LIST_NAME } from "../shared/constants";
import { getSP } from "../services/pnpjsConfig";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import {
  TextField,
  Dropdown,
  Selection,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
  DefaultButton,
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
import { IPnpService, PnpServices } from "../services/pnpservices";
import { Logger } from "@pnp/logging";

const stackTokens = { childrenGap: 50 };

const DelIcon: IIconProps = { iconName: "Delete" };
const ReadIcon: IIconProps = { iconName: "BulletedListText" };
const AddIcon: IIconProps = { iconName: "Add" };
const SaveIcon: IIconProps = { iconName: "Save" };

//Define Styles
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

const ddlBatchOptions: IDropdownOption[] = [
  { key: "Batch 1", text: "Batch 1" },
  { key: "Batch 2", text: "Batch 2" },
  { key: "Batch 3", text: "Batch 3" },
];

const ddlLevelOfKnowledgeOptions: IDropdownOption[] = [
  { key: "Beginner", text: "Beginner" },
  { key: "Intermediate", text: "Intermediate" },
  { key: "Expert", text: "Expert" },
];

const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};

export default class ReactPnpwp extends React.Component<
  IReactPnpwpProps,
  IReactPnpwpState
> {
  //private _services: PnpServices= null;
  private _selection: Selection;
  private _sp: PnpServices;

  constructor(props: IReactPnpwpProps, state: IReactPnpwpState) {
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
    console.log("React Constructor" + this._sp);

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
      this.setState({ ListItems: listItems, status: message });
      console.log(listItems);
    });
  }

  private async _createItem(): Promise<any> {
    await this._sp
      .CreateItem(this.props.listName, this.state.ListItem)
      .then((Id) => {
        this.callAndBindDetailsList("New Item Created With Id " + Id);
      });
  }
  private async _readList(): Promise<any> {
    await this.callAndBindDetailsList("Items Loaded Successfully");
  }
  private async _updateItem(): Promise<any> {
    await this._sp
      .UpdateItem(this.props.listName, this.state.ListItem.Id, {
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
        .DeleteItem(this.props.listName, this.state.ListItem.Id)
        .then(() => {
          this.setState({ status: "Item Deleted Successfully" });
        });
    } catch (e) {
      console.log(e);
    }
    alert("Item Deleted Successfully");
  }

  componentDidMount(): void {
    this.callAndBindDetailsList("Record Loaded");
  }

  public render(): React.ReactElement<IReactPnpwpProps> {
    const { description, context, listName } = this.props;

    return (
      <div>
        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <TextField
              label="User Name"
              placeholder="Please enter username here"
              value={this.state.ListItem.Title}
              onChange={(e, newValue) => {
                this.setState(
                  (state) => ((state.ListItem.Title = newValue), state)
                );
              }}
            />
            <TextField
              label="Email"
              placeholder="Please enter Email here"
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
              placeholder="Select Level Of Knowledge"
              label="Select  Level Of Knowledge"
              options={ddlLevelOfKnowledgeOptions}
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
        <div>
          <Stack horizontal tokens={stackTokens}>
            <PrimaryButton
              text="Create "
              iconProps={AddIcon}
              onClick={(e) => {
                this._createItem();
              }}
            />
            <PrimaryButton
              text="Read"
              iconProps={ReadIcon}
              onClick={(e) => {
                this._readList();
              }}
            />
            <PrimaryButton
              text="Update"
              iconProps={SaveIcon}
              onClick={(e) => this._updateItem()}
            />
            <PrimaryButton
              text="Delete"
              iconProps={DelIcon}
              onClick={(e) => this._deleteItem()}
            />
          </Stack>
        </div>
        <div id="divStatus">{this.state.status}</div>
        <DetailsList
          items={this.state.ListItems}
          columns={LIST_COLUMNS}
          setKey="Id"
          checkboxVisibility={CheckboxVisibility.onHover}
          selectionMode={SelectionMode.single}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          compact={true}
          selection={this._selection}
        />
      </div>
    );
  }
}
