import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { WEBURL, LISTID } from "../shared/constants";
import { IListItem } from "../models/IListItem";
export interface IGraphServices {
  CreateItem(itemObj: any): Promise<any>;
  getItems(): Promise<any>;
  updateItem(itemObj: any, itemId: number): Promise<any>;
  deleteItem(itemId: number): Promise<any>;
}

export class GraphServices implements IGraphServices {
  private _sp = null;

  constructor(context: WebPartContext) {
    this._sp = context.msGraphClientFactory;
  }

  public async CreateItem(itemObj: any): Promise<any> {
    try {
      let requestUrl;
      requestUrl = `/sites/${WEBURL}:/lists/${LISTID}/items`;
      let iar: any;
      await this._sp.getClient().then((client: MSGraphClient): void => {
        client.api(requestUrl).post(itemObj);
      });
    } catch (error) {
      Promise.reject(error);
      return error;
    }
  }

  public async getItems(): Promise<IListItem[]> {
    try {
      let requestUrl;
      requestUrl = `/sites/${WEBURL}:/lists/${LISTID}/items?expand=fields`;
      console.log(requestUrl);
      var listItems = [];
      await this._sp.getClient().then(async (client: MSGraphClient) => {
        client
          .api(requestUrl)
          .get()
          .then((response) => {
            console.log("Inside Promise block");
            console.log(response.value);
            var count = 0;

            response.value.forEach((element) => {
              console.log("New Item:", count, element.fields);
              count = count + 1;
              let { id, Title, Email, Batch, LevelOfKnowledge } =
                element.fields;
              var Id = Number(id);
              listItems.push({ Id, Title, Email, Batch, LevelOfKnowledge });
            });
          });
      });
      console.log("MyListItem-", listItems);
      return listItems;
    } catch (error) {
      Promise.reject(error);
      return error;
    }
  }
  public async updateItem(itemObj: any, itemId: number): Promise<any> {
    try {
      let requestUrl;
      requestUrl = `/sites/${WEBURL}:/lists/${LISTID}/items/${itemId}/fields`;
      await this._sp.getClient().then((client: MSGraphClient): void => {
        client.api(requestUrl).update(itemObj);
      });
      return;
    } catch (error) {
      Promise.reject(error);
      return error;
    }
  }

  public async deleteItem(itemId: number): Promise<any> {
    try {
      let requestUrl;
      requestUrl = `/sites/${WEBURL}:/lists/${LISTID}/items/${itemId}`;
      await this._sp.getClient().then((client: MSGraphClient): void => {
        client.api(requestUrl).delete();
      });
      return;
    } catch (error) {
      Promise.reject(error);
      return error;
    }
  }
}
