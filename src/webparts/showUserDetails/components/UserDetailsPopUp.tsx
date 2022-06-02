
import * as React from 'react';
import {UserDialogContentProps} from './UserDetailsPopUpProps';
import styles from './ShowUserDetails.module.scss';
import pnp from 'sp-pnp-js';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import * as ReactDOM from 'react-dom';
import { LazyLoadImage } from 'react-lazy-load-image-component';

export class UserDialogContent extends React.Component<UserDialogContentProps, {}> {

    constructor(props) {
      super(props);
      this.state = {
        profilePic: "",
        role: "",
        name: "",
        slogan: "",
        note: "",
      };
    }
    
  
  }
  
  export class UserDialog extends BaseDialog {
    public profilePic: string;
    public name: string;
    public role: string;
    public slogan: string;
    public note: string;
  
    public render(): void {
      var section = document.createElement('div');
      section.setAttribute("style",'align-items: center; width: 500px; height: auto; text-align: center; flex-direction: column; display: flex; align-items: center; ');
      
      var img  = document.createElement('img');
      img.setAttribute("src", this.profilePic);
      img.setAttribute("id", "profilePic");
      img.setAttribute("style", 'border-radius: 50%; width: 200px; height: 200px; object-fit: cover; margin-top: 30px; margin-bottom: 5px;');

      
      var name = document.createElement('span');
      name.innerHTML = this.name;
      name.setAttribute("style",'color: orange; font-size: 20px');


      var role = document.createElement('span');
      role.setAttribute("id","role");
      role.innerHTML = this.role;


      var note = document.createElement('div');
      note.setAttribute("id","note");
      note.setAttribute("style", 'margin: 30px;');
      note.innerHTML = this.note;

      var setText = document.createElement('span');
      setText.innerHTML = "Che cos'è Agic per me";
      setText.setAttribute("style",'color: orange; font-size: 20px');

      var slogan = document.createElement('div');
      slogan.setAttribute("id","slogan");
      slogan.setAttribute("style",'margin-inline: 30px; margin-bottom: 30px');
      slogan.innerHTML = this.slogan;

      section.append(img);
      section.append(name);
      section.append(role);
      section.append(note);
      if(slogan.innerText!=""){
        section.append(setText);
        section.append(slogan);}
      


      this.domElement.append(section);
      
      }  
      public getConfig(): IDialogConfiguration {
        return { isBlocking: false };
      }

      protected onAfterClose(): void {
        super.onAfterClose();
        this.domElement.childNodes[0].remove();
      }
    
      
        }
    