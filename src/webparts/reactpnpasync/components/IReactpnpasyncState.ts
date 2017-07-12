import { INewsItem } from "../interfaces";
export interface IReactpnpasyncState {
    items: INewsItem[];
    errors: string[];
    status: string;
}