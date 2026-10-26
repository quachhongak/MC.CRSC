import { ICrscPrompt } from '../models/ICrscPrompt';

export type IPromptSortOrder = 'newest' | 'oldest' | 'aToZ' | 'zToA';

export function filterAndSortPrompts(
  prompts: ICrscPrompt[],
  query: string,
  deliverable: string,
  category: string,
  sortOrder: IPromptSortOrder
): ICrscPrompt[] {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  const matchingPrompts = prompts.filter(prompt => {
    const matchesQuery = !normalizedQuery || [
      prompt.Deliverable,
      prompt.Title,
      prompt.AITool,
      prompt.RiskLevel,
      prompt.Category
    ].some(value => (value || '').toLocaleLowerCase().indexOf(normalizedQuery) !== -1);

    return matchesQuery &&
      (!deliverable || prompt.Deliverable === deliverable) &&
      (!category || prompt.Category === category);
  });

  return matchingPrompts.sort((left, right) => {
    switch (sortOrder) {
      case 'oldest':
        return new Date(left.Modified).getTime() - new Date(right.Modified).getTime();
      case 'aToZ':
        return left.Title.localeCompare(right.Title);
      case 'zToA':
        return right.Title.localeCompare(left.Title);
      case 'newest':
      default:
        return new Date(right.Modified).getTime() - new Date(left.Modified).getTime();
    }
  });
}

export function formatModifiedDate(value: string): string {
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return '';
  }

  const day = ('0' + date.getDate()).slice(-2);
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const year = String(date.getFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
}
