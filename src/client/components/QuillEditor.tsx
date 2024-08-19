import { IconButton, Label } from "@fluentui/react";
import { clone, set } from "lodash";
import * as React from "react";
import { useState } from "react";
// import * as Quill from "quill";
import { Control, useController } from "react-hook-form";
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css';
import { OutlookItem } from "../services/OutlookItem";

export interface IQuillComponentProps {
    label: string;
    html: string;
    outlookItem?: OutlookItem;
    placeholder?: string;
    required?: boolean;
    onChange?: (content: string) => void;
    taskTitle?: string;
}
export interface IQuillComponentState {
    editorHtml: string;
    mailChecked: boolean;
}
//export class QuillEditor extends React.Component<IQuillComponentProps, IQuillComponentState> {
export const QuillEditor = (props: IQuillComponentProps) => {
    const [editorHtml, setEditorHtml] = useState(props.html);
    const [mailChecked, setMailChecked] = useState(false);
    // constructor(props: IQuillComponentProps) {
    //     super(props);
    //     this.state = {
    //         editorHtml: props.html,
    //         mailChecked: false
    //     };
    //     this.handleChange = this.handleChange.bind(this);
    //     this.footerHandler = this.footerHandler.bind(this);
    //     this.taskTitleHandler = this.taskTitleHandler.bind(this);
    // }

    const CustomButton = () => <span title="insert email content" className="ms-Icon ms-Icon--Mail" />;

    /*
     * Custom toolbar component including insertStar button and dropdowns
     */
    const CustomToolbar = () => (
        <div id="toolbar">
            <select className="ql-header" defaultValue={""} onChange={e => e.persist()}>
                <option value="1" />
                <option value="2" />
                <option value="" />
            </select>
            <span className="ql-formats">
                <button className="ql-bold" />
                <button className="ql-italic" />
                <button className="ql-underline" />
                <button className="ql-strike" />
                <button className="ql-blockquote" />
            </span>
            <span className="ql-formats">
                <button className="ql-list" value="ordered" />
                <button className="ql-list" value="bullet" />
                <button className="ql-indent" value="-1" />
                <button className="ql-indent" value="+1" />
            </span>
            <span className="ql-formats">
                <button className="ql-link" />
                <button className="ql-image" />
                {/* <button className="ql-video" /> */}
                {/* <select className="ql-color">
                <option value="red" />
                <option value="green" />
                <option value="blue" />
                <option value="orange" />
                <option value="violet" />
                <option value="#d0d1d2" />
                <option selected />
            </select> */}
                <IconButton iconProps={{ iconName: 'mail' }}
                    disabled={props.outlookItem == null}
                    toggle
                    checked={mailChecked}
                    onClick={_footerHandler} />
                <IconButton iconProps={{ iconName: 'copy' }}
                    disabled={props.taskTitle == null || props.taskTitle == ""}
                    toggle
                    onClick={_taskTitleHandler} />
            </span>
        </div>
    );
    const _handleChange = (html: string) => {
        let newValue = "";
        if (html != "<p><br></p>")
            newValue = html;
        //console.log("newValue", newValue);
        setEditorHtml(newValue);
        if (props.onChange) {
            props.onChange(newValue);
        }
    }
    const _footerHandler = () => {
        //let newValue="<h3>welcome to the world</h3>";
        let newValue = props.outlookItem.Body;
        if (mailChecked) {
            newValue = "";
        }
        setEditorHtml(newValue);
        setMailChecked(!mailChecked);
        if (props.onChange) {
            props.onChange(newValue);
        }
    }
    const _taskTitleHandler = () => {
        let newValue = props.taskTitle;
        setEditorHtml(newValue);
        if (props.onChange) {
            props.onChange(newValue);
        }

    }

    /* 
   * Quill modules to attach to editor
   * See https://quilljs.com/docs/modules/ for complete options
   */
    const modules = {
        toolbar: {
            container: "#toolbar",
            // handlers: {
            //     footer: this.footerHandler
            // }
        },
        clipboard: {
            matchVisual: false,
        }
    };

    /* 
     * Quill editor formats
     * See https://quilljs.com/docs/formats/
     */
    const formats = [
        "header",
        "font",
        "size",
        "bold",
        "italic",
        "underline",
        "strike",
        "blockquote",
        "list",
        "bullet",
        "indent",
        "link",
        "image",
        "color"
    ];

    //   /* 
    //    * PropType validation
    //    */
    //   private propTypes = {
    //     placeholder: PropTypes.string
    //   };
    return (
        <div className="text-editor">
            <Label required={props.required}>{props.label}</Label>
            {/* <CustomToolbar /> */}
            {CustomToolbar()}
            <ReactQuill
                defaultValue={props.html}
                onChange={_handleChange}
                value={editorHtml}
                placeholder={props.placeholder}
                modules={modules}
                formats={formats}
                className="myQuill"
            //theme={"snow"} // pass false to use minimal theme
            />
        </div>
    );

}

