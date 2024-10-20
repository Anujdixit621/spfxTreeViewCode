import { ITreeItem } from "@pnp/spfx-controls-react/lib/TreeView";
import { IDropdownOption } from "office-ui-fabric-react";

// export interface IFolderTreeViewState {
//   folders: ITreeItem[];
//   loading: boolean;
//   selectedFolderItems: any[]; // To store the subfolders and files
//   error: string | null;
//   contentTypeOptions: any[];
//   contentOptions:string;
//   fieldsOptions:any[];
//   SelectedlibraryName:string;
// }

// Define a custom type for selected folder items (folders and files)


export interface IFolderTreeViewState {
 
  entityName:string;
  searchText:string;
  folders: ITreeItem[]; // The folder structure items
  loading: boolean;
  selectedFolderItems: IFolderOrFile[]; // Using the new custom type for folder/file items
  error: string ;
  contentTypeOptions: IDropdownOption[]; // Using IDropdownOption for dropdown options
  contentOptions: string;
  fieldsOptions: IDropdownOption[]; // Using IDropdownOption for field options in dropdown
  SelectedlibraryName: string;
}
interface IFolderOrFile {
  key: string;
  name: string | JSX.Element;
//  type: "Folder" | "File"; // To distinguish between folders and files
  ServerRelativeUrl?: string; // Optional, since folders may not have this
}

export interface IFileInfo {
  readonly "odata.id": string;
  CheckInComment: string;
  CheckOutType: number;
  ContentTag: string;
  CustomizedPageStatus: number;
  ETag: string;
  Exists: boolean;
  IrmEnabled: boolean;
  Length: string;
  Level: number;
  LinkingUri: string | null;
  LinkingUrl: string;
  ListId: string;
  MajorVersion: number;
  MinorVersion: number;
  Name: string;
  ServerRelativeUrl: string;
  SiteId: string;
  TimeCreated: string;
  TimeLastModified: string;
  Title: string | null;
  UIVersion: number;
  UIVersionLabel: string;
  UniqueId: string;
  WebId: string;
}