import { ICrscPrompt } from '../src/webparts/crscPrompts/models/ICrscPrompt';
import MockSharePointService from '../src/webparts/crscPrompts/services/MockSharePointService';
import { filterAndSortPrompts, formatModifiedDate } from '../src/webparts/crscPrompts/utils/promptUtils';

const prompts: ICrscPrompt[] = [
  {
    Id: 1,
    Title: 'Write a report',
    Input: 'Summarize the source material.',
    AITool: 'Copilot',
    Category: 'Writing',
    Deliverable: 'Report',
    RiskLevel: 'Low',
    SourceInputsRequired: '',
    LimitationsandConstraints: '',
    AdditionalInformation: '',
    Modified: '2026-01-02T00:00:00Z'
  },
  {
    Id: 2,
    Title: 'Plan a project',
    Input: 'Create a project plan.',
    AITool: 'ChatGPT',
    Category: 'Planning',
    Deliverable: 'Plan',
    RiskLevel: 'Medium',
    SourceInputsRequired: '',
    LimitationsandConstraints: '',
    AdditionalInformation: '',
    Modified: '2026-01-03T00:00:00Z'
  }
];

describe('filterAndSortPrompts', () => {
  it('searches the configured fields without regard to case', () => {
    expect(filterAndSortPrompts(prompts, 'COPILOT', '', '', 'newest').map(prompt => prompt.Id)).toEqual([1]);
    expect(filterAndSortPrompts(prompts, 'medium', '', '', 'newest').map(prompt => prompt.Id)).toEqual([2]);
  });

  it('applies deliverable and category filters together', () => {
    expect(filterAndSortPrompts(prompts, '', 'Plan', 'Planning', 'newest').map(prompt => prompt.Id)).toEqual([2]);
  });

  it('sorts newest first by default and supports the requested alternatives', () => {
    expect(filterAndSortPrompts(prompts, '', '', '', 'newest').map(prompt => prompt.Id)).toEqual([2, 1]);
    expect(filterAndSortPrompts(prompts, '', '', '', 'oldest').map(prompt => prompt.Id)).toEqual([1, 2]);
    expect(filterAndSortPrompts(prompts, '', '', '', 'aToZ').map(prompt => prompt.Id)).toEqual([2, 1]);
    expect(filterAndSortPrompts(prompts, '', '', '', 'zToA').map(prompt => prompt.Id)).toEqual([1, 2]);
  });
});

describe('formatModifiedDate', () => {
  it('formats dates as DD-MM-YY and returns an empty value for invalid dates', () => {
    expect(formatModifiedDate('2026-01-02T00:00:00Z')).toMatch(/^\d{2}-\d{2}-26$/);
    expect(formatModifiedDate('not a date')).toBe('');
  });
});

describe('MockSharePointService', () => {
  it('provides local list and prompt data without contacting SharePoint', async () => {
    const service = new MockSharePointService();
    await expect(service.getLists()).resolves.toEqual([{ id: 'mock-list', title: 'Prompt Library' }]);
    await expect(service.getPrompts('mock-list')).resolves.toHaveLength(1);
  });
});
