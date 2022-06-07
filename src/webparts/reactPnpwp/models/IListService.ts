export interface IListService {
    //getDocumentLibrary(): Promise<JSON>;
    //getDocumentLibraryWithPnPJS(): Promise<JSON>;
    createItem(payload:any):Promise<JSON>;
    readItem():Promise<JSON>;
    updateItem(id:string):Promise<JSON>;
    deleteItem(id:string):Promise<JSON>;
}