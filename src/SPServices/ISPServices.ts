import {  SearchResults } from '@pnp/sp/search';


export interface ISPServices {

    searchUsers2(): Promise<SearchResults>;
    searchUsers(searchString: string, searchFirstName: boolean): Promise<SearchResults>;
    searchUsersNew(searchString: string, srchQry: string, isInitialSearch: boolean): Promise<SearchResults>;

}
