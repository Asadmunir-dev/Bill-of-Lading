export interface NotifyParty {
  id: string;
  name: string;
  address: string;
}

export interface ContainerItem {
  id: string;
  containerNo: string;
  netWeightKg: number;
  grossWeightKg: number;
  packageCount: number;
}

export interface BillOfLadingRecord {
  id: string;
  userId: string;
  userEmail?: string;
  headerTitle: string;
  blNumber: string;
  bookingRef: string;
  carrierName: string;
  shipperText: string;
  referenceNo: string;
  consigneeText: string;
  freightPayableAt: string;
  formENumber: string;
  vesselAndVoy: string;
  placeOfDelivery: string;
  portOfLoading: string;
  portOfDischarge: string;
  marksAndNumbers: string;
  cargoDescription: string;
  grossWeight: string;
  grossWeightUnit: string;
  nettWeight: string;
  nettWeightUnit: string;
  totalPackagesCount: string;
  totalPackagesUnit: string;
  measurement: string;
  containerSummaryLine: string;
  totalContainersPackages: string;
  movement: string;
  numberOfOriginals: string;
  placeAndDateOfIssue: string;
  specialClauses: string[];
  notifyParties: NotifyParty[];
  containers: ContainerItem[];
  savedAt?: string;
  updatedAt?: string;
}

export interface UserProfile {
  uid: string;
  email: string;
  displayName?: string;
  photoURL?: string;
  defaultShipper?: string;
  updatedAt?: string;
}
