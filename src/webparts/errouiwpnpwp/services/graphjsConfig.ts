import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp and pnp logging system

import { graphfi, GraphFI, SPFx } from "@pnp/graph";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";

var _sp: GraphFI = null;

export const getSP = (context?: WebPartContext): GraphFI => {
  if (_sp === null && context != null) {
    _sp = graphfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));

    console.log("GetSP Inside", _sp);
  }
  console.log("GetSP Outside", _sp);
  return _sp;
};
