import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

import TagInput from "./TagInput";
import { saveTagData, getTagData, TagData } from "./TagData";


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
  tags: TagData[];

  isAddingTag: boolean;
  newTag: string;
  newText: string;
  // tags: { tag: string; text: string }[];
}


export default class App extends React.Component<AppProps, AppState> {
  constructor(props: AppProps, context: any) {
    super(props, context);
    this.state = {
      listItems: [],
      tags: [],

      isAddingTag: false,
      newTag: "",
      newText: "",
      // tags: [],
    };
  }

  componentDidMount() {
    const tags = getTagData();
    this.setState({tags});
  }

  handleSaveTagData = (tag: string, text: string) => {
    saveTagData(tag, text);
    const tagData = getTagData();
    this.setState({ tags:tagData });
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
    const { tags } = this.state;
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
        <div>
        <h1>Tag Data:</h1>
        <ul>
          {tags.map((data, index) => (
            <li key={index}>
              <strong>Tag:</strong> {data.tag}, <strong>Text:</strong> {data.text}
            </li>
          ))}
        </ul>
        <h2>Add Tag and Text</h2>
        <TagInput onSave={this.handleSaveTagData} />
      </div>
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
