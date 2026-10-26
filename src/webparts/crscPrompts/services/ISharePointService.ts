import { ICrscPrompt, IListOption } from '../models/ICrscPrompt';

export interface ISharePointService {
  getLists(): Promise<IListOption[]>;
  getPrompts(listId: string): Promise<ICrscPrompt[]>;
}
