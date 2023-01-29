import IAdmissionItem from '../models/IAdmissionItem';
export interface IPhcpadmissionsProps {
  webparttitle: string;
  itens: IAdmissionItem[];
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
