import { ISharePointService } from '../services/ISharePointService';

export interface ICrscPromptsProps {
  listId?: string;
  itemsPerPage: number;
  modalMinWidth: number;
  clearAllButtonColor: string;
  cardBorderColor: string;
  cardTitleColor: string;
  viewButtonTextColor: string;
  viewButtonBackgroundColor: string;
  service: ISharePointService;
  onConfigure: () => void;
}
