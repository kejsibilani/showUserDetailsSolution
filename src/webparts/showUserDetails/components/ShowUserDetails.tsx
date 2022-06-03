import * as React from 'react';
import styles from './ShowUserDetails.module.scss';
import { IShowUserDetailsProps } from './IShowUserDetailsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as pnp from 'sp-pnp-js';
import {UserDialog} from './UserDetailsPopUp';
import { forEach } from 'lodash';
import { addWeeks } from 'office-ui-fabric-react';

export default class ShowUserDetails extends React.Component<IShowUserDetailsProps, {
  allItemsWithPics: any[];
  level: string;
  items: any[];
  listName: string;
  context: string;
  errorDisplay: string;
  tableDisplay: string;
  type: string;
  spinnerDisplay: string;
}> {
  
  constructor(props){
    super(props);
    this.state = {
      context: this.props.context.pageContext.web.absoluteUrl,
      level: "",
      listName: "",
      items: [],
      allItemsWithPics: [],
      errorDisplay: "none",
      tableDisplay: "flex",
      type: "",
      spinnerDisplay: "",
    };
    this.popUp = this.popUp.bind(this);
  }

  public componentDidMount(): void {  

    try{
      this.showAll();
      }
      catch {
        this.setState({
                errorDisplay: "",
                tableDisplay: "none"
              });
      }
    
  }




  private async showAll() {
    var itemsAll: any[] = [];
    var listName = this.props.listName;
    var level = this.props.level;
    var files: any[] = [];
    try {await pnp.sp.web.lists.getByTitle(listName).items.orderBy("Order0", true).filter("Level eq '" + level + "'").get().then(items => 
      { items.forEach(item => {itemsAll.push(item);});});

    
    for (var i = 0; i< itemsAll.length; i++){
     await pnp.sp.web.lists.getByTitle(listName).items.getById(itemsAll[i].ID).file.expand("File/Name").get().then(fname => {
        var img = this.state.context + "/" + listName.replace('&',"") + "/" + fname.Name;
        itemsAll[i].img = img;
      });
    }

    this.setState({
      items: itemsAll,
      spinnerDisplay: "none"
    });}

    catch 
      {
        this.setState({
                spinnerDisplay: "none",
                errorDisplay: "",
                tableDisplay: "none"
              });
      }
    
    }
   
     

    private popUp(index): React.MouseEventHandler<HTMLButtonElement>{
      const dialog: UserDialog = new UserDialog();
      dialog.profilePic = this.state.items[index].img;
      dialog.name = this.state.items[index].Title;
      dialog.slogan = this.state.items[index].Slogan;
      dialog.role = this.state.items[index].Role;
      dialog.note = this.state.items[index].Note;
      dialog.show();
      return;
    } 

  public render(): React.ReactElement<IShowUserDetailsProps> {
    const {
      level,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.showUserDetails} ${hasTeamsContext ? styles.teams : ''}`} >
             <h2 style={{textAlign: 'center'}} id="levelDisplay"></h2>  
        <br/>
        <div id="showUsers" style={{display:this.state.tableDisplay, overflowX: 'auto', columnCount: '3', flexWrap: 'wrap'}}> 
          {
            this.state.items.map((item, index) => 
            <>
            <section className = "column" style={{width: '230px', height: '230px', alignItems: 'center', textAlign: 'center'}} >
              <img id="profilePic" onClick={()=>this.popUp(index)} src={item["img"]} style={{borderRadius: '50%', width: '150px', height: '150px', objectFit: 'cover'}}/>
            <div id="name">{item["Title"]}</div>
            <div>{item["Role"]}</div>
            </section>
            </>
            )
          }
        </div>
        <div style={{display:this.state.errorDisplay}}>Input Error</div>
        <div style={{display:this.state.spinnerDisplay}}> <div className={`${styles.spinner}`}></div> </div> 
      </section>
    );
  }
}
