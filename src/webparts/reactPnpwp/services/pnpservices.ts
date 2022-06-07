import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { AadTokenProviderFactory } from "@microsoft/sp-http";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import { IItemAddResult } from "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "./pnpjsConfig";
import { IList } from "@pnp/sp/lists";
import { IListItem } from "../models/IListItem";

export interface IPnpService {
  CreateItem(listName: string, itemObj: any): Promise<any>;
  getItems(listName: string, columns: string[]): Promise<any[]>;
  UpdateItem(listName: string, itemId: number, itemObj: any): Promise<any>;
  DeleteItem(listName: string, itemId: number): Promise<any>;
}

export class PnpServices implements IPnpService {
  //public static readonly serviceKey: ServiceKey<IPnpService> = ServiceKey.create<IPnpService>('SPFx:PnpServices', PnpServices);
  private _sp;

  //private context:WebPartContext;

  constructor(context: WebPartContext) {
    this._sp = getSP(context);
    console.log("From PnpServices:" + this._sp);
  }

  public async getItems(listName: String): Promise<any> {
    //, columns: string[]
    try {
      const items: any[] = await this._sp.web.lists
        .getByTitle(listName)
        .items();

      console.log(items);

      return items;
    } catch (err) {
      Promise.reject(err);
      return err;
    }
  }
  public async CreateItem(listName: string, itemObj: any): Promise<any> {
    try {
      const iar: IItemAddResult = await this._sp.web.lists
        .getByTitle(listName)
        .items.add(itemObj);

      return iar.data.Id;
    } catch (err) {
      Promise.reject(err);
      return err;
    }
  }

  public async UpdateItem(
    listName: string,
    itemId: number,
    itemObj: any
  ): Promise<any> {
    try {
      console.log(JSON.stringify(itemObj));
      const list = this._sp.web.lists.getByTitle(listName);

      const i = await list.items.getById(itemId).update(itemObj);

      console.log(i);
      return itemId;
    } catch (err) {
      Promise.reject(err);
      return err;
    }
  }

  public async DeleteItem(listName: string, itemId: number): Promise<any> {
    try {
      const list = this._sp.web.lists.getByTitle(listName);

      const i = await list.items.getById(itemId).delete();
      return;
    } catch (err) {
      Promise.reject(err);
      return err;
    }
  }
}
