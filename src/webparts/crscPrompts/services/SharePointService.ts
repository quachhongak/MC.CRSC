import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { ICrscPrompt, IListOption } from '../models/ICrscPrompt';
import { ISharePointService } from './ISharePointService';

const promptFields: string[] = [
  'Id',
  'Title',
  'Input',
  'AITool',
  'Category',
  'Deliverable',
  'RiskLevel',
  'SourceInputsRequired',
  'LimitationsandConstraints',
  'AdditionalInformation',
  'Modified'
];

export default class SharePointService implements ISharePointService {
  public constructor(private readonly sp: SPFI) {}

  public async getLists(): Promise<IListOption[]> {
    const lists = await this.sp.web.lists
      .select('Id', 'Title', 'Hidden')
      .filter('Hidden eq false')
      .orderBy('Title')
      .top(200)();

    return lists.map(list => ({ id: list.Id, title: list.Title }));
  }

  public async getPrompts(listId: string): Promise<ICrscPrompt[]> {
    const prompts: ICrscPrompt[] = [];
    const query = this.sp.web.lists.getById(listId).items
      .select(...promptFields)
      .filter("Status eq 'Approved'")
      .orderBy('Modified', false)
      .top(100);
    const pages = query[Symbol.asyncIterator]();
    let page = await pages.next();

    while (!page.done) {
      prompts.push(...page.value as ICrscPrompt[]);
      page = await pages.next();
    }

    return prompts;
  }
}
