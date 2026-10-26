import { ICrscPrompt, IListOption } from '../models/ICrscPrompt';
import { ISharePointService } from './ISharePointService';

export default class MockSharePointService implements ISharePointService {
  public getLists(): Promise<IListOption[]> {
    return Promise.resolve([{ id: 'mock-list', title: 'Prompt Library' }]);
  }

  public getPrompts(_listId: string): Promise<ICrscPrompt[]> {
    return Promise.resolve([{
      Id: 1,
      Title: 'Example prompt',
      Input: 'Describe the task to complete.',
      AITool: 'Copilot',
      Category: 'General',
      Deliverable: 'Draft',
      RiskLevel: 'Low',
      SourceInputsRequired: '',
      LimitationsandConstraints: '',
      AdditionalInformation: '',
      Modified: new Date().toISOString()
    }]);
  }
}
