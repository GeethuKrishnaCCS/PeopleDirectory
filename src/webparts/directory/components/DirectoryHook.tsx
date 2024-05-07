/* eslint-disable @typescript-eslint/no-floating-promises*/
import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from "./Directory.module.scss";
import { PersonaCard } from "./PersonaCard/PersonaCard";
import { spservices } from "../../../SPServices/spservices";
import { IDirectoryState } from "./IDirectoryState";
import * as strings from "DirectoryWebPartStrings";
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType, SearchBox, Icon, Label,
  Dropdown, IDropdownOption, IconButton, TooltipHost, IDropdownStyles, ISearchBoxStyles, Stack, IStackTokens
} from "@fluentui/react";
import { debounce } from "throttle-debounce";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import { ISPServices } from "../../../SPServices/ISPServices";


import { IDirectoryProps } from './IDirectoryProps';
import Paging from './Pagination/Paging';


const wrapStackTokens: IStackTokens = { childrenGap: 30 };

const DirectoryHook: React.FC<IDirectoryProps> = (props) => {
  const _services: ISPServices = new spservices(props.context);
  const [az, setaz] = useState<string[]>([]);
  const [alphaKey, setalphaKey] = useState<string>('');
  const [state, setstate] = useState<IDirectoryState>({
    users: [],
    isLoading: true,
    errorMessage: "",
    hasError: false,
    indexSelectedKey: "",
    searchString: "",
    searchText: "",
  });
  // const xxx = React.useRef(false);
  const initialState = React.useRef(state.users);
  const orderOptions: IDropdownOption[] = [
    { key: "FirstName", text: "First Name" },
    { key: "LastName", text: "Last Name" },
    { key: "Department", text: "Department" },
    { key: "Location", text: "Location" },
    { key: "JobTitle", text: "Job Title" }
  ];
  const color = props.context.microsoftTeams ? "white" : "";
  // Paging
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [pagedItems, setPagedItems] = useState<any[]>([]);
  const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [filterSelectDropdownOptions, setFilterSelectDropdownOptions] = useState<any[]>([]);
  const [selectedFilters, setSelectedFilter] = useState<any[]>([]);


  const refinerString = props.filterSettings.refiners;
  const refiners = refinerString && refinerString.length > 0 ? refinerString.split(',') : [];

  const _onPageUpdate = async (pageno?: number): Promise<void> => {
    console.log("alphabet", az);
    const currentPge = (pageno) ? pageno : currentPage;
    const startItem = ((currentPge - 1) * pageSize);
    const endItem = currentPge * pageSize;
    const filItems = state.users.slice(startItem, endItem);
    console.log("users", state.users);

    setCurrentPage(currentPge);
    setPagedItems(filItems);
  };

  useEffect(() => {
    if (state.users.length === 0 || initialState.current.length !== 0) return
    initialState.current = state.users;
  }, [state.users]);

  const filterUsers = (users: any[]) => {
    // filtering logic to the users according to the filter settings
    return users.filter((user: { Department: any; JobTitle: any; WorkPhone: any; WorkEmail: any; BaseOfficeLocation: any; }) => {
      const filterSettings = props.filterSettings;
      if (filterSettings.hideUsersWithoutDept && !user.Department) {
        return false;
      }
      if (filterSettings.hideUsersWithoutJobTitle && !user.JobTitle) {
        return false;
      }
      if (filterSettings.hideUsersWithoutPhone && !user.WorkPhone) {
        return false;
      }
      if (filterSettings.hideUsersWithoutEmail && !user.WorkEmail) {
        return false;
      }
      if (filterSettings.hideUsersWithoutLocation && !user.BaseOfficeLocation) {
        return false;
      }
      return true;
    });
  };

  const diretoryGrid =
    pagedItems && pagedItems.length > 0
      ? pagedItems
        .map((user: any, i) => {
          return (
            <PersonaCard
              context={props.context}
              key={"PersonaCard" + i}
              profileProperties={{
                DisplayName: user.PreferredName,
                Title: props.cardSettings.showUserJobTitle && user.JobTitle,
                PictureUrl: props.cardSettings.showUserPhoto && user.PictureURL,
                Email: user.WorkEmail.toLowerCase(),
                Department: props.cardSettings.showUserDept && user.Department,
                WorkPhone: props.cardSettings.showUserPhone && user.WorkPhone,
                Location: props.cardSettings.showUserLocation && user.BaseOfficeLocation
              }}
            />
          );
        })
      : [];
  const _loadAlphabets = (): void => {
    const alphabets: string[] = [];
    for (let i = 65; i < 91; i++) {
      alphabets.push(
        String.fromCharCode(i)
      );
    }
    setaz(alphabets);
  };

  const _searchByAlphabets = async (initialSearch: boolean): Promise<void> => {
    console.log("_searchByAlphabets", initialSearch)
    setstate({ ...state, isLoading: true, searchText: '' });
    let users = null;
    if (initialSearch) {
      users = await _services.searchUsersNew('', '', true);
    } else {
      users = await _services.searchUsersNew(`${alphaKey}`, '', true);
    }
    // if (initialSearch) {
    //   if (props.searchFirstName)
    //     //   users = await _services.searchUsers2();
    //     // else users = await _services.searchUsers2();
    //     users = await _services.searchUsersNew('', `FirstName:a*`, false);
    //   else users = await _services.searchUsersNew('', '', true);
    // } else {
    //   if (props.searchFirstName)
    //     users = await _services.searchUsersNew('', `FirstName:${alphaKey}*`, false);
    //   else users = await _services.searchUsersNew(`${alphaKey}`, '', true);
    //   // users = await _services.searchUsers2();
    //   // else users = await _services.searchUsers2();
    // }
    // filtering logic to the fetched users according to the filter settings
    const filteredUsers = filterUsers(users?.PrimarySearchResults.sort() || []);
    // Sort filteredUsers array alphabetically by firstname
    filteredUsers.sort((a, b) => {
      const nameA = a.FirstName.toUpperCase();
      const nameB = b.FirstName.toUpperCase();
      if (nameA < nameB) {
        return -1;
      }
      if (nameA > nameB) {
        return 1;
      }
      return 0; // names must be equal
    });

    setstate({
      ...state,
      searchText: '',
      // indexSelectedKey: initialSearch ? 'A' : state.indexSelectedKey,
      users: filteredUsers,
      isLoading: false,
      errorMessage: "",
      hasError: false
    });
  };

  const _searchUsers = async (searchText: string): Promise<void> => {
    console.log("_searchUsers", searchText)
    try {
      setstate({ ...state, searchText: searchText, isLoading: true });
      if (searchText.length > 0) {
        const searchProps: string[] = props.searchProps && props.searchProps.length > 0 ?
          props.searchProps.split(',') : ['FirstName', 'LastName', 'WorkEmail', 'Department'];
        let qryText = '';
        const finalSearchText: string = searchText ? searchText.replace(/ /g, '+') : searchText;
        if (props.clearTextSearchProps) {
          const tmpCTProps: string[] = props.clearTextSearchProps.indexOf(',') >= 0 ? props.clearTextSearchProps.split(',') : [props.clearTextSearchProps];
          if (tmpCTProps.length > 0) {
            searchProps.map((srchprop, index) => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const ctPresent: any[] = tmpCTProps.filter((o) => { return o.toLowerCase() === srchprop.toLowerCase(); });
              if (ctPresent.length > 0) {
                if (index === searchProps.length - 1) {
                  qryText += `${srchprop}:${searchText}*`;
                } else qryText += `${srchprop}:${searchText}* OR `;
              } else {
                if (index === searchProps.length - 1) {
                  qryText += `${srchprop}:${finalSearchText}*`;
                } else qryText += `${srchprop}:${finalSearchText}* OR `;
              }
            });
          } else {
            searchProps.map((srchprop, index) => {
              if (index === searchProps.length - 1)
                qryText += `${srchprop}:${finalSearchText}*`;
              else qryText += `${srchprop}:${finalSearchText}* OR `;
            });
          }
        } else {
          searchProps.map((srchprop, index) => {
            if (index === searchProps.length - 1)
              qryText += `${srchprop}:${finalSearchText}*`;
            else qryText += `${srchprop}:${finalSearchText}* OR `;
          });
        }
        console.log(qryText);
        const users = await _services.searchUsersNew('', qryText, false);
        // const users = await _services.searchUsers2();
        setstate({
          ...state,
          searchText: searchText,
          // indexSelectedKey: '0',
          users:
            users && users.PrimarySearchResults
              ? users.PrimarySearchResults
              : null,
          isLoading: false,
          errorMessage: "",
          hasError: false
        });
        setalphaKey('0');
      } else {
        setstate({ ...state, searchText: '' });
        await _searchByAlphabets(true);
      }
    } catch (err) {
      setstate({ ...state, errorMessage: err.message, hasError: true });
    }
  };
  const _debouncesearchUsers = debounce(500, _searchUsers);

  const _searchBoxChanged = (newvalue: string): void => {
    setCurrentPage(1);
    _debouncesearchUsers(newvalue);
  };


  const _sortPeople = async (sortField: string): Promise<void> => {
    let _users = [...state.users];
    _users = _users.sort((a: any, b: any) => {

      switch (sortField) {

        // Sort by Location
        case "Location":
          if ((a.BaseOfficeLocation || "").toUpperCase() < (b.BaseOfficeLocation || "").toUpperCase()) {
            return -1;
          }
          if ((a.BaseOfficeLocation || "").toUpperCase() > (b.BaseOfficeLocation || "").toUpperCase()) {
            return 1;
          }
          return 0;

          break;

        default:
          if ((a[sortField] || "").toUpperCase() < (b[sortField] || "").toUpperCase()) {
            return -1;
          }
          if ((a[sortField] || "").toUpperCase() > (b[sortField] || "").toUpperCase()) {
            return 1;
          }
          return 0;

          break;
      }
    });
    setstate({ ...state, users: _users, searchString: sortField });
  };
  //write a function for a button which sorts in ascending and descending order
  const _changeSortDirection = async (): Promise<void> => {
    let _users = [...state.users];
    const reversedArray = _users.reverse();
    setstate({ ...state, users: reversedArray });

  }
  const _filterPeople = async (filterConditions: any): Promise<void> => {
    let _users = [...initialState.current];

    // Apply each filter condition
    Object.keys(filterConditions).forEach(filterField => {
      const filterBy = filterConditions[filterField];
      // Filter users based on the current filter condition
      _users = _users.filter(user => {
        // eslint-disable-next-line
        return user.hasOwnProperty(filterField) && user[filterField] === filterBy;
      });
    });

    // Update state with filtered and sorted users
    setstate({ ...state, users: _users });
  };

  useEffect(() => {
    setPageSize(props.pageSize);
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    if (state.users) { _onPageUpdate() }
  }, [state.users, props.pageSize]);

  useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    if (alphaKey.length > 0 && alphaKey !== "0") _searchByAlphabets(false);
  }, [alphaKey]);

  useEffect(() => {
    _loadAlphabets();
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    _searchByAlphabets(true);
  }, [props]);

  useEffect(() => {
    let _users = [...initialState.current];
    let dropdownOptionsByRefiner: any = {};

    refiners.forEach((refiner) => {
      // Filter users based on selected filters 
      // Object.keys(selectedFilters).forEach((filterKey) => {
      //   _users = _users.filter((user) => user[filterKey] === selectedFilters[filterKey as any]);
      // });

      const filterOptions = _users.reduce((options, user) => {
        if (Object.prototype.hasOwnProperty.call(user, refiner)) {
          const value = user[refiner];
          if (value !== null && !options.includes(value)) {
            options.push(value);
          }
        }
        return options;
      }, []);

      const dropdownOptions = filterOptions.map((option: any) => ({
        key: option,
        text: option.toString(),
      }));

      dropdownOptionsByRefiner[refiner] = dropdownOptions;
    });
    setFilterSelectDropdownOptions(dropdownOptionsByRefiner);

  }, [state.users, selectedFilters]);

  const handleClearFiltersClick = () => {
    _searchByAlphabets(true);
    setSelectedFilter([]);
  };

  const dropdownStyles: Partial<IDropdownStyles> = {
    title: {
      borderRadius: 100,
      borderWidth: 2,
      borderColor: "#DDD9D9",
      height: 30,
      lineHeight: 30,
      paddingLeft: 20,
      color: "white",
      fontSize: 14,
      backgroundColor: "black",
    },
    caretDownWrapper: {
      height: 30,
      lineHeight: 30,
      float: "right",
    },
    caretDown: {
      fontSize: 17,
      color: "white",
      width: "auto"
    },
    label: {
      color: "#323338",
      fontSize: 1
    },
  };
  const searchboxStyles: Partial<ISearchBoxStyles> = {
    root: {
      borderRadius: 100,
      borderWidth: 2,
      borderColor: "#DDD9D9",
      height: 48,
      lineHeight: 41,
      paddingLeft: 20,
      color: "#666666",
      fontSize: 16,
    }
  };
  return (
    <div className={styles.directory}>
      <WebPartTitle displayMode={props.displayMode} title={props.title}
        updateProperty={props.updateProperty} />
      <div className={styles.searchBox}>
        <SearchBox placeholder={strings.SearchPlaceHolder} className={styles.searchTextBox}
          onSearch={_searchUsers}
          value={state.searchText}
          onChange={(ev, newVal) => _searchBoxChanged(newVal)}
          styles={searchboxStyles} />
        <div className={styles.dropDownSortBy}>
          {refiners.length > 0 && <div className={styles.filterDiv}>
            <div key={filterSelectDropdownOptions.length} style={{
              display: 'flex', alignItems: 'center',
              overflow: 'hidden',
              flexWrap: 'wrap',
              gap: '10px'
            }} >
              {refiners.map((refiner, index) => (
                <Dropdown
                  key={refiner}
                  placeholder={`${refiner}`}
                  options={filterSelectDropdownOptions[refiner as keyof typeof filterSelectDropdownOptions]}
                  selectedKey={selectedFilters[refiner as any] || ""}
                  calloutProps={{ calloutWidth: undefined, calloutMinWidth: 100, calloutMaxWidth: 440 }}
                  onChange={(ev, value) => {
                    const filterConditions = {
                      ...selectedFilters, // existing selected filter conditions
                      [refiner]: value.key.toString()
                    };
                    setSelectedFilter(filterConditions)

                    _filterPeople(filterConditions);
                  }}
                  styles={dropdownStyles}
                />
              ))}
              <div>
                <TooltipHost
                  content="Clear All Filters"
                >
                  <IconButton iconProps={{ iconName: 'Refresh' }} aria-label="ClearAllFilters" onClick={handleClearFiltersClick} />
                </TooltipHost>
              </div>
            </div>
            <div className={styles.dropDownSortBy}>
            </div>
          </div>}
          {state.users.length > 0 && <div className={styles.sortDiv}>
            <div key={filterSelectDropdownOptions.length} style={{ display: 'flex', alignItems: 'center' }}>
              <Dropdown
                placeholder={strings.DropDownPlaceHolderMessage}
                // label={strings.DropDownPlaceLabelMessage}
                options={orderOptions}
                selectedKey={state.searchString}
                onChange={(ev, value) => {
                  // eslint-disable-next-line @typescript-eslint/no-floating-promises
                  _sortPeople(value.key.toString());
                }}
                styles={dropdownStyles}
              // styles={{ dropdown: { width: 200 } }}
              />
              <div>
                <TooltipHost
                  content="Sort in Ascending or Descending order"
                >
                  <IconButton iconProps={{ iconName: 'Sort' }} aria-label="SortUporDown" onClick={_changeSortDirection} />
                </TooltipHost>
              </div>
            </div>
          </div>}
        </div>
        <div>
        </div>
      </div>
      {state.isLoading ? (
        <div style={{ marginTop: '10px' }}>
          <Spinner size={SpinnerSize.large} label={strings.LoadingText} />
        </div>
      ) : (
        <>
          {state.hasError ? (
            <div style={{ marginTop: '10px' }}>
              <MessageBar messageBarType={MessageBarType.error}>
                {state.errorMessage}
              </MessageBar>
            </div>
          ) : (
            <>
              {!pagedItems || pagedItems.length === 0 ? (
                <div className={styles.noUsers}>
                  <Icon
                    iconName={"ProfileSearch"}
                    style={{ fontSize: "54px", color: color }}
                  />
                  <Label>
                    <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>
                      {strings.DirectoryMessage}
                    </span>
                  </Label>
                </div>
              ) : (
                <>
                  {/* <div style={{ width: '100%', display: 'inline-block' }}>
                    <Paging
                      totalItems={state.users.length}
                      itemsCountPerPage={pageSize}
                      onPageUpdate={_onPageUpdate}
                      currentPage={currentPage} />
                  </div> */}

                  <Stack horizontal
                    horizontalAlign={props.useSpaceBetween ? "space-between" : "center"}
                    tokens={wrapStackTokens}
                    wrap>
                    {diretoryGrid}
                  </Stack>
                 
               
                {/* <div className={styles.directoryGrid} style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(233px, 1fr))', gap: '16px' }}>
                  {diretoryGrid}
                </div> */}


                  <div style={{ width: '100%', display: 'inline-block' }}>
                    <Paging
                      totalItems={state.users.length}
                      itemsCountPerPage={pageSize}
                      onPageUpdate={_onPageUpdate}
                      currentPage={currentPage}
                      pageRange={props.pageRange}
                    />
                  </div>
                  <div>
                  </div>
                  {/* for debugging purposes */}
                  {/*  <div>
                    <h2>Total Users: {state.users.length}</h2>
                    <ul>
                      {state.users.map((user: { FirstName: boolean | React.ReactChild | React.ReactFragment | React.ReactPortal; LastName: boolean | React.ReactChild | React.ReactFragment | React.ReactPortal; Department: boolean | React.ReactChild | React.ReactFragment | React.ReactPortal; }, index: React.Key) => (
                        <li key={index}>
                          <strong>Name:</strong> {user.FirstName} {user.LastName} <br />
                          <strong>Department:</strong> {user.Department}
                        </li>
                      ))}
                    </ul>
                  </div> */}
                </>
              )}
            </>
          )}
        </>
      )}
    </div>
  );
};

export default DirectoryHook;
