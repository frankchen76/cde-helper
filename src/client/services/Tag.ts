import { ITag } from "@fluentui/react";

export class Tag {
    // private _id: string;
    // private _name: string;
    // private _url: string;

    // public get Id() { return this._id; }
    // public get Name() { return this._name; }
    // public get Url() { return this._url; }

    // constructor(row: any) {
    //     this._id = row["id"];
    //     this._name = row["name"];
    //     this._url = row["url"];
    // }
    public id: string;
    public name: string;
    public url: string;
    public static createTagFromResponse(response: any): Tag {
        let ret = new Tag();
        ret.id = response["id"];
        ret.name = response["name"];
        ret.url = response["url"];
        return ret;
    }
    public toITag(): ITag {
        return { key: this.id, name: this.name };
    }
}

export class TagCollection {
    // private _tags: Tag[];

    // public get Tags() { return this._tags; }

    // constructor(response: any) {
    //     if (response["value"]) {
    //         this._tags = [];
    //         response["value"].forEach(row => {
    //             this._tags.push(new Tag(row));
    //         });
    //     }
    // }
    constructor(public items: Tag[]) {

    }
    public static createTagsFromResponse(response: any): TagCollection {
        let ret = null;
        if (response && response.value) {
            ret = new TagCollection([]);
            response.value.forEach(itemResponse => {
                ret.items.push(Tag.createTagFromResponse(itemResponse));
            })
        }
        return ret;
    }
    public toITags(): ITag[] {
        let ret: ITag[] = [];
        if (this.items) {
            ret = this.items.map(t => t.toITag());
        }
        return ret;
    }
    public getITagsFromTagString(tags: string): ITag[] {
        let ret: ITag[] = [];
        if (tags && tags != "") {
            const tagStringArray = tags.indexOf(";") == -1 ? [tags] : tags.split(";").map(t => t.trim());
            const existTags = this.items.filter(t => tagStringArray.indexOf(t.name) != -1);
            ret = existTags ? existTags.map(t => t.toITag()) : [];
        }
        return ret;
    }
}