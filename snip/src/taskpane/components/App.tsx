import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

// export interface AppState {
//   listItems: HeroListItem[];
// }
export interface AppState {
  listItems: HeroListItem[];
  isAddingTag: boolean;
  newTag: string;
  newText: string;
  tags: { tag: string; text: string }[];
}


export default class App extends React.Component<AppProps, AppState> {
  constructor(props: AppProps, context: any) {
    super(props, context);
    this.state = {
      listItems: [],
      isAddingTag: false,
      newTag: "",
      newText: "",
      tags: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  handleSaveTagData = (tag: string, text: string) => {
    saveTagData(tag, text);
    const tagData = getTagData();
    this.setState({ tagData });
  };
  
  click = async () => {
    if (this.state.isAddingTag) {
      // Save the new tag and text
      const { newTag, newText } = this.state;
      const newTags = [...this.state.tags, { tag: newTag, text: newText }];
      this.setState({ tags: newTags, isAddingTag: false, newTag: "", newText: "" });
    } else {
      this.setState({ isAddingTag: true });
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Bienvenido" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          {this.state.isAddingTag ? (
            <div>
              <input
                type="text"
                placeholder="Enter tag"
                value={this.state.newTag}
                onChange={(e) => this.setState({ newTag: e.target.value })}
              />
              <input
                type="text"
                placeholder="Enter text"
                value={this.state.newText}
                onChange={(e) => this.setState({ newText: e.target.value })}
              />
              <button onClick={this.click}>Save Tag</button>
            </div>
          ) : (
            <p className="ms-font-l">
              Modify the source files, then click <b>Run</b>.
            </p>
          )}
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            {this.state.isAddingTag ? "Cancel" : "Agregar tag y texto"}
          </DefaultButton>
        </HeroList>
        <div className="data-grid">
          {this.state.tags.map((tagObj, index) => (
            <div key={index}>
              <div className="tag">{tagObj.tag}</div>
              <div className="text">{tagObj.text}</div>
            </div>
          ))}
        </div>
      </div>
    );
  }
}
