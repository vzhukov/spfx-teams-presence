import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPresence } from "../../../model/IPresence";

/**
 * Properties of the component
 */
export interface IUserPresenceProps {
  context: WebPartContext;
}

/**
 * State of the component
 */
export interface IUserPresenceState{
  userUPN?: string;
  userId?: string;
  presence?: IPresence;
}