import * as React from "react";
import { TextField } from 'office-ui-fabric-react/lib/TextField';

export interface ITextFieldProps {
    multiline: boolean;
    textFieldChanged?: (newValue: string) => void;
    textFieldBlured?: () => void;
    inputText?: string;
    textAreaFieldLimit: number;
  }

  export interface ITextFieldState
  extends React.ComponentState,
  ITextFieldProps {}
 
  export class QuickReplyText extends React.Component<
  ITextFieldProps,
  ITextFieldState
> {
  constructor(props: ITextFieldProps) {
    super(props);
    
    this.state = {
      multiline: props.multiline,  
      inputText: props.inputText,
      textAreaFieldLimit : props.textAreaFieldLimit
    };
  }

  public componentWillReceiveProps(newProps: ITextFieldProps): void {
    this.setState(newProps);
  }

render(): JSX.Element {
    return (
        <TextField
        multiline={this.state.multiline}
        onChange={this._onChange}
        list="options"
        id="textField"
        onBlur={this._onBlur}
        value={this.state.inputText}     
      />
    );
  }

  private _onBlur = (ev : any) : void => 
  {
    if (this.props.textFieldBlured) {
      this.props.textFieldBlured();
    }
  }

  private _onChange = (ev: any, newText: string): void => {
    const newMultiline = newText.length > this.state.textAreaFieldLimit;
    this.setState({inputText: newText});

    if (newMultiline !== this.state.multiline) {
      this.setState({ multiline: newMultiline });
    }

    if (this.props.textFieldChanged) {
      this.props.textFieldChanged(newText);
    }
  };
}


