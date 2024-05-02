import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
export interface IDirectoryProps {
  title: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  searchFirstName: boolean;
  updateProperty: (value: string) => void;
  searchProps?: string;
  clearTextSearchProps?: string;
  pageSize?: number;
  useSpaceBetween?: boolean;
  cardSettings: cardSettings;
  filterSettings: filterSettings;
}

interface cardSettings {
  showUserPhoto: boolean,
  showUserDept: boolean,
  showUserJobTitle: boolean,
  showUserPhone: boolean,
  showUserEmail: boolean,
  showUserLocation: boolean
}

interface filterSettings {
  hideUsersWithoutDept: boolean,
  hideUsersWithoutJobTitle: boolean,
  hideUsersWithoutPhone: boolean,
  hideUsersWithoutEmail: boolean,
  hideUsersWithoutLocation: boolean,
  refiners: string
}