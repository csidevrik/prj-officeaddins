export interface TagData {
    tag: string;
    text: string;
  }
  
  export function saveTagData(tag: string, text: string): void {
    const existingData = localStorage.getItem("tagData");
    let data: TagData[] = [];
  
    if (existingData) {
      data = JSON.parse(existingData);
    }
  
    data.push({ tag, text });
    localStorage.setItem("tagData", JSON.stringify(data));
  }
  
  export function getTagData(): TagData[] {
    const existingData = localStorage.getItem("tagData");
  
    if (existingData) {
      return JSON.parse(existingData);
    }
  
    return [];
  }
  