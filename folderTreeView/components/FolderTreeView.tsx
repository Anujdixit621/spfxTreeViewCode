import * as React from "react";
import { IFolderTreeViewProps } from "./IFolderTreeViewProps";
import { IFolderTreeViewState } from "./IFolderTreeViewState";
import {   sp } from "@pnp/sp/presets/all";
import { ISearchQuery } from "@pnp/sp/search";
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  mergeStyleSets,
  TextField,
  PrimaryButton,
  
} from "office-ui-fabric-react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { TreeView, ITreeItem } from "@pnp/spfx-controls-react/lib/TreeView";
//import { Callout } from "@fluentui/react";
//import { IFolder } from "@pnp/spfx-controls-react";

const classNames = mergeStyleSets({
  container: {
    display: "flex",
    padding: "20px",
  },
  treeViewSection: {
    width: "25%",
    marginRight: "20px",
    backgroundColor: "#cff4ec",
    padding: "10px",
    borderRadius: "5px",
  },
  detailsSection: {
    width: "50%",
    backgroundColor: "#ffffff",
    padding: "10px",
    borderRadius: "5px",
    boxShadow: "0px 4px 6px rgba(0, 0, 0, 0.1)",
  },
  dropdownSection: {
    width: "20%",
    backgroundColor: "#f3f2f1",
    padding: "10px",
    borderRadius: "5px",
  },
  title: {
    fontWeight: "bold",
    marginBottom: "10px",
  },
});
interface SearchResult {
  // Define the structure of the items you expect to receive
  [key: string]: any;
}
export default class FolderTreeView extends React.Component<IFolderTreeViewProps, IFolderTreeViewState> {
  constructor(props: IFolderTreeViewProps) {
    super(props);

    this.state = {
      
      entityName:"",
      searchText:"",
      folders: [],
      selectedFolderItems: [],
      loading: true,
      error: "",
      contentTypeOptions: [],
      contentOptions: "",
      fieldsOptions: [],
      SelectedlibraryName: "",
    };

    this.handleTextInputChange = this.handleTextInputChange.bind(this);
    this.handleEntityInputChange = this.handleEntityInputChange.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    const siteName = this.props.siteName;
    try {
      
      const rootFolderUrl = "/sites/IT/IT";

    
    this.getSubfolderNested(rootFolderUrl).then(subFolders => {
  console.info("All folders and subfolders:", subFolders);
});
      this.handleSearch();
      this.getAllLibrariesWithSubFolders().then((libraries) => {
        console.log("All Libraries with their folders: ", libraries);
    });
      this.getCustomContentTypeFields("Path_Event")
      const libraryName ="Document Library";
      const getContentTypes = await this.getContentTypes(libraryName);  // 12oct
      console.info(`Content types of library "${libraryName}":`, getContentTypes);
      const _getAllFilesAndMetadata = await this._getAllFilesAndMetadata(libraryName);// 12oct
      console.info(`All files with its metadata when any folder select "${libraryName}":`, _getAllFilesAndMetadata);
      const searchText = "HITACHI"; // Replace with the text you want to search for
      const results = await this._searchInLibrary(libraryName, searchText);
      console.log(`Search results for "${searchText}" in library "${libraryName}":`, results);
      const documentLibraries = await this._getDocumentLibraries(siteName);
      const libraryFolders: ITreeItem[] = [];
      for (const library of documentLibraries) {
        const rootFolders = await this._getFoldersFromLibrary(library.Title); // Fetch top-level folders
        const libraryItem: ITreeItem = {
          key: library.Id,
          label: library.Title,
          children: rootFolders, // Initialize with root-level folders
        };
        libraryFolders.push(libraryItem);
      }

      this.setState({ folders: libraryFolders, loading: false });
    } catch (error) {
      this.setState({ loading: false, error: error.message });
    }
  }

  // Fetch document libraries
  private async _getDocumentLibraries(siteName: string): Promise<{ Title: string; Id: string }[]> {
    sp.setup({
      sp: { baseUrl: siteName },
    });
    const documentLibraries = await sp.web.lists
      .filter("BaseTemplate eq 101")
      .select("Title", "Id")
      .get();
    return documentLibraries;
  }
///////////Nested Subfolder 



 // Fetch folders from a document library
private async _getFoldersFromLibrary(libraryTitle: string): Promise<ITreeItem[]> {
  const items: ITreeItem[] = [];
  const subfolders = await sp.web.lists
    .getByTitle(libraryTitle)
    .items.filter("FSObjType eq 1") // Filter only folders (FSObjType = 1)
    .select("FileLeafRef", "FileRef")
    .get();

  for (const folder of subfolders) {
    const folderItem: ITreeItem = {
      key: folder.FileRef,
      label: folder.FileLeafRef,
      children: [], // Ensure children is initialized as an empty array
      data: { hasChildren: true }, // Mark that this item might have children
    };
    items.push(folderItem);
  }

  return items;
}
//// Final Result function

private handleSearch = async () => {
  
  let filters: string[] = [];

  let customFields: { key: string, value: string }[] = [
    { key: 'Title', value: 'Document1' },
    { key: 'Author', value: 'John Doe' }
  ];
  if (customFields.length > 0) {
    customFields.forEach(field => {
      filters.push(`${field.key} eq '${field.value}'`);
    });
  }
let freeText = "Alstom";
let radioOption = "AllFields";

  // Free text search based on radio selection
  if (freeText.length > 0) {
    if (radioOption === 'AllFields') {
      filters.push(`substringof('${freeText}', Title) or substringof('${freeText}', Author/Title)`);
    } else if (radioOption === 'DocumentText') {
      // Call SharePoint search API with document text query
      const results = await sp.search({ Querytext: freeText });
      console.info(`if Radio button of ALL Feilds selected Result :${results}`)
    } else if (radioOption === 'EntityName') {
      filters.push(`EntityName eq '${freeText}'`);
    }
  }
  let startDate = null;
  let endDate = null;
  // Date range filters
  if (startDate && endDate) {
    filters.push(`Created ge datetime'${startDate}' and Created le datetime'${endDate}'`);
  }
let contentType = ""; 
  // Content type filter
  if (contentType.length > 0) {
    filters.push(`ContentType eq '${contentType}'`);
  }

  // Final query to be executed
  const query = filters.join(' and ');
  console.log('Final query:', query);

  // Execute the query
  const results = await sp.web.lists.getByTitle('LibraryName')
    .items.filter(query)
    .get();

  console.log('Search results:', results);
};


/////
// Fetch subfolders and files recursively
private async _getSubFolders(folderUrl: string): Promise<ITreeItem[]> {
  const items: ITreeItem[] = [];
  const folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
  const subFolders = await folder.folders.select("Name", "ServerRelativeUrl").get();

  for (const subFolder of subFolders) {
    const subFolderItem: ITreeItem = {
      key: subFolder.ServerRelativeUrl,
      label: subFolder.Name,
      children: [], // Ensure children is initialized as an empty array
      data: { hasChildren: true },
    };
    items.push(subFolderItem);
  }

  return items;
}

private handleFolderSelection = async (item: ITreeItem, libraryName: string): Promise<void> => {
  
  const folderUrl = item.key; // Use the folder URL (key) directly

  if (!folderUrl) return;

  // If the folder has not been expanded before, fetch its subfolders
  if (item.data?.hasChildren && (!item.children || item.children.length === 0)) {
    const subFolders = await this._getSubFolders(folderUrl);
    item.children = subFolders; // Update the item's children with subfolders
  }

  // Fetch files only for the selected folder, not the entire library
  const files: any[] = await this._getFilesInSpecificFolder(folderUrl);

  // Construct folder and file items to bind to the DetailsList
  const folderItems = [
    ...(item.children ? item.children.map((subFolder) => ({
      key: subFolder.key,
      name: subFolder.label,
      type: "Folder",
    })) : []), // Use a fallback to an empty array if item.children is undefined
    ...files.map((file) => ({
      key: file.ServerRelativeUrl,
      name: file.Name,
      type: "File",
    })),
  ];

  // Update the state with the selected folder's items (folders and files)
  this.setState({ selectedFolderItems: folderItems, SelectedlibraryName: libraryName });

  // Fetch content types specific to the selected library
  const contentTypeOptions = await this.getContentTypes(libraryName);
  this.setState({ contentTypeOptions });
};


  private async _getFilesInSpecificFolder(folderUrl: string): Promise<{ Name: string; ServerRelativeUrl: string }[]> {
    const folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
    const files = await folder.files.select("Name", "ServerRelativeUrl", "FileLeafRef", "FileRef").get();
    return files;
  }
  // Fetch content types for the dropdown
  private async getContentTypes(libraryName: string): Promise<IDropdownOption[]> {
    const siteName = this.props.siteName;
    sp.setup({
      sp: { baseUrl: siteName },
    });
    const contentTypes = await sp.web.lists.getByTitle(libraryName).contentTypes.get();
    return contentTypes.map((ct) => ({
      key: ct.Id.StringValue,
      text: ct.Name,
    }));
  }
  ///
private async getSubfolderNested(folderUrl: string){
try{
  debugger;
  const foldersAccumulator = [];
  const siteName = this.props.siteName;
  sp.setup({
    sp: { baseUrl: siteName },
  });
  const subFolders = await sp.web
      .getFolderByServerRelativeUrl(folderUrl)
      .folders
      .select("Name", "ServerRelativeUrl") // Only get the fields you need
      .get();
      console.info(`All sub folder ${subFolders}`);
      foldersAccumulator.push(...subFolders);
      for (const folder of subFolders) {
        await getAllSubFolders(folder.ServerRelativeUrl, foldersAccumulator);
      }
  
      return foldersAccumulator;
}
catch{

}
}
//////16 oct night
private async getAllLibrariesWithSubFolders(){
try{
  debugger;
  const siteName = this.props.siteName;
  sp.setup({
    sp: { baseUrl: siteName },
  });
  const libraries = await sp.web.lists.filter("BaseTemplate eq 101 and Hidden eq false").get();

  const librariesWithFolders = await Promise.all(libraries.map(async (library) => {
      const rootFolders = await sp.web.lists.getById(library.Id).rootFolder.folders.get();
      const allFolders = this.getSubFolders(rootFolders);
      return {
          library: library.Title,
          folders: allFolders
      };
  }));

  console.log("Libraries and Folders:", librariesWithFolders);
  return librariesWithFolders;
} 
catch(error){
  return [];
}
}

private async getSubFolders(folders: any[]){
  try{
    const subfoldersPromises = folders.map(async (folder: { ServerRelativeUrl: string; subfolders: Promise<any> | never[]; }) => {
      // Fetch subfolders of the current folder
      const subfolders = await sp.web.getFolderByServerRelativeUrl(folder.ServerRelativeUrl).folders.get();

      if (subfolders.length > 0) {
          folder.subfolders = this.getSubFolders(subfolders);
      } else {
          folder.subfolders = [];
      }
      return folder;
  });
  return await Promise.all(subfoldersPromises);
  } 
  catch{

  }
}
// async function getSubFolders(folders: any[]): Promise<any[]> {
//     const subfoldersPromises = folders.map(async (folder) => {
//         // Fetch subfolders of the current folder
//         const subfolders = await sp.web.getFolderByServerRelativeUrl(folder.ServerRelativeUrl).folders.get();

//         if (subfolders.length > 0) {
//             folder.subfolders = await getSubFolders(subfolders);
//         } else {
//             folder.subfolders = [];
//         }
//         return folder;
//     });

//     return await Promise.all(subfoldersPromises);
// }
/////


  /// get all files from select folder 
  private async _getAllFilesAndMetadata(libraryName: string): Promise<any[]> {
    try {
      
      const library = sp.web.lists.getByTitle(libraryName);
      const files = await library.items.get();
      //console.info(`Files with its metadata "${files}"`)
      console.log(...files);
      return files;
    } catch (error) {
      console.error(`Error fetching files and metadata from library "${libraryName}": `, error);
      return [];
    }
  }
  /// if ALl fields radio button click 
  
private async _getFieldsFromLibrary(libraryName: string): Promise<{ InternalName: string; TypeAsString: string }[]> {
  try {
    const library = sp.web.lists.getByTitle(libraryName);
    const fields = await library.fields.select("InternalName", "TypeAsString").get(); 
    console.info("Old Lib",fields);
   const filterfield = await sp.web.lists.getByTitle(libraryName).fields
  .select("InternalName", "TypeAsString", "Filterable", "ReadOnlyField", "Hidden","Title")
  .filter("Hidden eq false and ReadOnlyField eq false and Filterable eq true")
  .get();
     const filterValue = "Alstom";
     const filterFieldNames = filterfield.map(field => `${field.InternalName} eq '${filterValue}'`).join(' or ');
     console.info("New lib",filterfield);
     console.info("Files",filterFieldNames);
     const libraryTestLIB = sp.web.lists.getByTitle(libraryName).fields.select("InternalName", "TypeAsString").get();
     console.info("New lib",libraryTestLIB);
    // Only fetch visible fields
     const TaxNomyFields =fields.filter(field => field.TypeAsString === "TaxonomyFieldType");
     console.info(`fetching TaxNomyFields from library "${libraryName}":`,TaxNomyFields)
     const textFields = fields.filter(field => field.TypeAsString === "Text");
     console.info(`fetching All fields from library "${libraryName}":`,textFields)
     return fields;
  } catch (error) {
    console.error(`Error fetching fields from library "${libraryName}":`, error);
    return [];
  }
}

private async _searchInLibrary(libraryName: string, searchText: string): Promise<any[]> {
  try {
    
    const fields = await this._getFieldsFromLibrary(libraryName); //12oct
    console.info(`Fileds name of "${libraryName}"`,fields.length);
    const TaxNomyFields =fields.filter(field => field.TypeAsString === "TaxonomyFieldType");
    console.info(`fetching All fields from library "${libraryName}":`,TaxNomyFields.length)
    const textFields = fields.filter(field => field.TypeAsString === "Text");
    const filterConditions = textFields.filter(field=>field.InternalName.indexOf("_")).map(field => `substringof('${searchText}', ${field.InternalName})`).join(" or ");
    const library = sp.web.lists.getByTitle(libraryName);
    const textFieldResults: SearchResult[]  = await library.items.filter(filterConditions).get();
    let camlQuery = this._buildTaxonomyCAMLQuery(TaxNomyFields, searchText);
    let taxonomyFieldResults: SearchResult[] = [];
    if (camlQuery) {
      taxonomyFieldResults = await library.getItemsByCAMLQuery({
        ViewXml: camlQuery,
      });
    }
    const allResults = [...textFieldResults, ...taxonomyFieldResults];
    return allResults;
  } catch (error) {
    console.error(`Error searching in library "${libraryName}":`, error);
    return error;
  }
}
///////CAML Query Nested API
// Helper function to create dynamic CAML query for taxonomy fields
private _buildTaxonomyCAMLQuery(fields: any[], searchText: string): string {
  const buildNestedOr = (conditions: string[]): string => {
    if (conditions.length === 1) {
      return conditions[0]; // Return the last condition if only one left
    } else if (conditions.length === 2) {
      return `<Or>${conditions[0]}${conditions[1]}</Or>`; // Wrap two conditions
    } else {
      const firstCondition = conditions.shift();
      return `<Or>${firstCondition}${buildNestedOr(conditions)}</Or>`; // Recursively build <Or> conditions
    }
  };
  const conditions = fields.map(field => `
    <Eq>
      <FieldRef Name='${field.InternalName}' />
      <Value Type='TaxonomyFieldType'>${searchText}</Value>
    </Eq>
  `);
  return `
    <View>
      <Query>
        <Where>
          ${buildNestedOr(conditions)}
        </Where>
      </Query>
    </View>
  `;
}

/////// Start Free Text
handleTextInputChange = (event: React.FormEvent<HTMLInputElement>, newValue?: string): void => {
  this.setState({ searchText: newValue || '' });
};
handleEntityInputChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
  this.setState({ entityName: newValue || '' });
}

// Function to perform the free text search when the button is clicked
freeTextSearch = async (): Promise<void> => {
  const { searchText } = this.state;

  debugger;
  if (searchText.trim() === '') {
    alert("Please enter text for search!");  // Alert if no text is entered
    return;
  }
  console.log('Performing search with text:', searchText);
  const libraryName="Path_Event";
  const siteURl = "https://pathinfotech365.sharepoint.com/sites/IT/";
  const Path =`${siteURl}${libraryName}/*`;
  console.info(`'${searchText} AND Path:${Path}'`);
  const searchPath =`${siteURl}${libraryName}/*`;
 // const queryhard =https://pathinfotech365.sharepoint.com/_api/search/query?querytext=%27Alstom%20AND%20Path:%22https://pathinfotech365.sharepoint.com/sites/IT/Path%22%27
  const searchQuery: ISearchQuery = {
  Querytext:  `stringof("${searchText}") AND Path:"${searchPath}"`.replace(/\\/g, ''),
   // Querytext: `stringof("${searchText}") AND Path:"${searchPath}"`,
    RowLimit: 50,  // Number of search results
    SelectProperties: [
      "Title",   // Retrieve document title
      "Path",    // Retrieve the document's URL path
      "FileExtension",  // Retrieve the document type (extension)
      "Author",  // Retrieve author information
      "Created", // Retrieve created date
      "Modified" // Retrieve modified date
    ],
    TrimDuplicates: true  // Remove duplicate search results
  };
  try {
    const searchResults = await sp.search(searchQuery); 
    console.log("Search results:", searchResults.PrimarySearchResults);
  } catch (error) {
    console.error("Error performing search:", error);
    alert("Error performing search. Please try again.");
  }
  //this.performSearch(searchText);
};
  ///////
  // if select entity name radio button 

  callEntityNameSearch = async(): Promise<void>=>{
    const {entityName} = this.state;
    const libraryName = "Path_Event"; // Specify the document library name
    const filterQuery = `substringof('${entityName}', FileLeafRef)`; // Filter by FileRef field containing entityName
    const library = sp.web.lists.getByTitle(libraryName);
    const fields = await library.fields.select("InternalName", "TypeAsString").filter("Hidden eq false").get();
    const excludedFieldNames = [
      'Editor',
      'CheckoutUser',
      'ItemChildCount',
      'FolderChildCount',
      'AppAuthor',
      'AppEditor',
      'ParentVersionString',
      'ParentLeafName'
    ];
    const fieldNames = fields
  .filter(field => field.InternalName.indexOf("_") === -1) // Keep fields that don't have '_'
  .map(field => field.InternalName) // Map to InternalName
  .filter(field => excludedFieldNames.indexOf(field) === -1); // Exclude specific fields
    //const fieldNames = fields.filter(field=>field.InternalName.indexOf("_")).map(field => field.InternalName);
    const searchResults = await sp.web.lists
      .getByTitle(libraryName) // Replace with your document library name
      .items.filter(filterQuery) // Filter items where FileRef contains entityName
      .select(...fieldNames,"Author/Title") // Specify fields to select
      .expand("Author") // Expand user fields like Author
      .top(50) // Limit the number of results
      .get();
      console.info("Search Results of entity Name:", searchResults);
  }

///////
///////Date range filter 
callDateRangeSearch = async () => {

const dateString = "2024-10-14T04:05:06Z";
// Option 1: Using JavaScript Date object
const dateObj = new Date(dateString);
const formattedDate = ('0' + dateObj.getDate()).slice(-2) + '/' + 
                      ('0' + (dateObj.getMonth() + 1)).slice(-2) + '/' + 
                      dateObj.getFullYear();


console.log(formattedDate); // Output: 08/10/2024
const startDate1="2024-10-14T00:00:00Z";
const endDate1="2024-10-14T23:59:59Z";
var formattedDate1="Created"
  //const filterQuery = `${selectedDateType} ge '${startDate}' and ${selectedDateType} le '${endDate}'`;
  const filterQuery = `${formattedDate1} ge datetime'${startDate1}' and ${formattedDate1} le datetime'${endDate1}'`;

  console.log('Date Filter Query:', filterQuery);

  // Replace with your document library name
  const libraryName = "Path_Event"; 

  // Use this query in your PnP JS filter
  const searchResults = await sp.web.lists
    .getByTitle(libraryName)
    .items
    .filter(filterQuery)
    .select('Title', 'FileRef', 'Author/Title', 'Created', 'Modified')  // Specify the fields to retrieve
    .expand('Author')  // Expand Author field
    .top(50)
    .get();

  console.log('Filtered Results by Date:', searchResults);
  // You can now process the search results (e.g., display them in your UI)
}
/// get content type filed 
//  private getFieldsByContentTypeNameFromLibrary = async(libraryTitle:string)=>{
//   try{
//     const siteName = this.props.siteName;
//     sp.setup({
//       sp: { baseUrl: siteName },
//     });
//     const contentTypes = await sp.web.lists.getByTitle(libraryTitle).contentTypes
//       .select("Name", "Id", "FieldLinks", "FieldLinks/Name")
//       .expand("FieldLinks")
//       .get();
//       console.info(...contentTypes); 
//   }
//   catch{}
// }


private async getCustomContentTypeFields(libraryName: string): Promise<void> {
  const siteName = this.props.siteName;
  sp.setup({
    sp: { baseUrl: siteName },
  });

  try {
    
    // Get content types from the library
    const contentTypes = await sp.web.lists.getByTitle(libraryName).contentTypes.get();

    // Loop through each content type to get their fields
    for (const ct of contentTypes) {
      // Get fields for the current content type
      const fields = await sp.web.lists.getByTitle(libraryName).contentTypes.getById(ct.Id.StringValue).fields.get();

      // Filter out system-defined fields
      const userDefinedFields = fields.filter(field => !field.FromBaseType);

      // Log the user-defined fields to the console
      console.log(`Fields for Content Type: ${ct.Name}`);
      userDefinedFields.forEach(field => {
        console.log(`Field Title: ${field.Title}, Internal Name: ${field.InternalName}`);
      });
    }
  } catch (error) {
    console.error("Error fetching custom fields:", error);
  }
}

/////
  handleSubmit(): void {
    const { searchText, entityName } = this.state;
    
    if (searchText) {
      this.freeTextSearch();  // Call free text search function
    } else if (entityName) {
      this.callEntityNameSearch(); // Call entity name search function
    } 
    else if (searchText=="" && entityName=="")
    {
      this.callDateRangeSearch();
    }
    // else {
    //   console.log('Please enter a value to search.');
    // }
  }

  /////
  private handleContentTypeChange = async (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): Promise<void> => {

    if (option) {
      const contentType = option.text as string;
      const items = await sp.web.lists
        .getByTitle(this.state.SelectedlibraryName)
        .items.filter(`ContentType eq '${contentType}'`)
        .select("Title", "Id", "FileLeafRef", "FileRef", "ContentType/Name")
        .expand("ContentType")
        .get();
      const mappedItems = items.map((file) => ({
        key: file.FileRef, // Use FileRef as the key
        name: (
          <a href={`${"https://pathinfotech365.sharepoint.com"}${file.FileRef}`} target="_blank" rel="noopener noreferrer">
            {file.FileLeafRef}
          </a>
        ),
        type: "File", // Set type to "File"
      }));
      this.setState({ selectedFolderItems: mappedItems });
      console.log("Fetched items:", mappedItems);
    }
  };

  public render(): React.ReactElement<IFolderTreeViewProps> {
    const { folders, selectedFolderItems, loading, error, contentTypeOptions } = this.state;
    const columns: IColumn[] = [
      { key: "column1", name: "Name", fieldName: "name", minWidth: 100, isResizable: true },
      { key: "column2", name: "Type", fieldName: "type", minWidth: 50, isResizable: true },
    ];

    if (loading) {
      return <Spinner size={SpinnerSize.large} label="Loading folders..." />;
    }

    if (error) {
      return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
    }

    return (
      <div className={classNames.container}>
        <div className={classNames.treeViewSection}>
          <div className={classNames.title}>Folders</div>
                     <TreeView
            items={folders}
           defaultExpanded={false}
          // selectionMode="single"
             showCheckboxes={false}
            treeItemActionsDisplayMode={1}
          expandToSelected={false}
          onSelect={(items: ITreeItem[]) => this.handleFolderSelection(items[0], items[0].label)}
            />
        </div>
        <div className={classNames.dropdownSection}>
          <div className={classNames.title}>Content Types</div>
          <Dropdown
            placeholder="Select Content Type"
            options={contentTypeOptions}
            onChange={this.handleContentTypeChange}
          />
           <TextField 
    placeholder="Enter some text"
    onChange={this.handleTextInputChange}  // Handle text input change event
    value={this.state.searchText}
  />
      <TextField 
          placeholder="Enter entity name" 
          onChange={this.handleEntityInputChange}  // Handle entity name input change event
          value={this.state.entityName}
        />
       
      
  {/* Button for any action */}
  <PrimaryButton
    text="Submit"
    onClick={this.handleSubmit}  // Handle button click event
  />
        </div>
        <div className={classNames.detailsSection}>
          <div className={classNames.title}>Files and Subfolders</div>
          <DetailsList
            items={selectedFolderItems}
            columns={columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
          />
        </div>
      </div>
    );
  }
}

async function getAllSubFolders(folderUrl: string,foldersAccumulator: import("@pnp/sp/folders/types").IFolderInfo[]) {
  try {
    const subFolders = await sp.web
      .getFolderByServerRelativeUrl(folderUrl)
      .folders
      .select("Name", "ServerRelativeUrl")
      .get();
    foldersAccumulator.push(...subFolders);
    for (const folder of subFolders) {
      await this.getAllSubFolders(folder.ServerRelativeUrl, foldersAccumulator);
    }
  } catch (error) {
    console.error(`Error retrieving subfolders from ${folderUrl}: `, error);
  }
}



