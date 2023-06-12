// import React, { useState } from "react";
import * as React from "react";

interface TagInputProps {
  onSave: (tag: string, text: string) => void;
}

const TagInput: React.FC<TagInputProps> = ({ onSave }) => {
  const [tag, setTag] = React.useState("");
  const [text, setText] = React.useState("");

  const handleSave = () => {
    onSave(tag, text);
    setTag("");
    setText("");
  };

  return (
    <div>
      <input
        type="text"
        value={tag}
        onChange={(e) => setTag(e.target.value)}
        placeholder="Enter tag"
      />
      <input
        type="text"
        value={text}
        onChange={(e) => setText(e.target.value)}
        placeholder="Enter text"
      />
      <button onClick={handleSave}>Save</button>
    </div>
  );
};

export default TagInput;
