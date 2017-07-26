import * as React from "react";
import * as ReactDOM from "react-dom";
import {action, observable, computed} from "mobx";
import {observer} from "mobx-react";

class Info {
    @observable private _userName:string;

    @computed public get UserName():string { return this._userName; }

    @action
    public updateName(userName:string): void {
        this._userName = userName;
    }
}

let info: Info = new Info();

Office.initialize = () => {
    info.updateName(Office.context.mailbox.userProfile.displayName);
}


@observer
class HelloWorld extends React.Component<{}, {}> {

    public render(): JSX.Element {
        return <div> Hello {info.UserName}! </div>;
    }
}

ReactDOM.render(
    (<HelloWorld />),
    document.getElementById("app")
);