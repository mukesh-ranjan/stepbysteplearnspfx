import {IListItem} from "../models/IListItem"

export interface IReactPnpwpState{
    status:string;
    ListItem:IListItem;
    ListItems:IListItem[];
}