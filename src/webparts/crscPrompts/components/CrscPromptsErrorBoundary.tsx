import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';
import { MessageBar, MessageBarType } from '@fluentui/react';
import * as strings from 'CrscPromptsWebPartStrings';

interface ICrscPromptsErrorBoundaryState {
  hasError: boolean;
}

export default class CrscPromptsErrorBoundary extends React.Component<React.PropsWithChildren<{}>, ICrscPromptsErrorBoundaryState> {
  public state: ICrscPromptsErrorBoundaryState = { hasError: false };

  public static getDerivedStateFromError(): ICrscPromptsErrorBoundaryState {
    return { hasError: true };
  }

  public componentDidCatch(error: Error): void {
    Log.error('CrscPrompts', error);
  }

  public render(): React.ReactNode {
    if (this.state.hasError) {
      return <MessageBar messageBarType={MessageBarType.error}>{strings.ComponentError}</MessageBar>;
    }

    return this.props.children;
  }
}
