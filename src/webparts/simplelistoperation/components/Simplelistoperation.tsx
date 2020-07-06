import * as React from 'react';
import styles from './Simplelistoperation.module.scss';
import { ISimplelistoperationProps } from './ISimplelistoperationProps';
import { ISimpleListOperationsState, IListItem } from './ISimpleListOperationsState';
import { TextField, DefaultButton, PrimaryButton, Stack, IStackTokens, IIconProps } from 'office-ui-fabric-react/lib/';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { escape } from '@microsoft/sp-lodash-subset';


// Sp Pnp Setup
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";

const stackTokens: IStackTokens= { childrenGap: 40};
const DelIcon : IIconProps = { iconName: 'Delete' };
const ClearIcon : IIconProps = { iconName: 'Clear' };
const AddIcon : IIconProps = { iconName: 'Add' };

export default class Simplelistoperation extends React.Component<ISimplelistoperationProps, ISimpleListOperationsState> {
  constructor (props: ISimplelistoperationProps, state: ISimpleListOperationsState){
    super(props);
    this.state = {
      addText: '',
      updateText: [],
      
     
    };

    if (Environment.type === EnvironmentType.SharePoint) {
      this._getListItems();
    }
    else if (Environment.type === EnvironmentType.Local) {
      // return (<div>Whoops! you are using local host...</div>);
    }
  }

  public render(): React.ReactElement<ISimplelistoperationProps> {
    return (
      <div className={ styles.simplelistoperation }>
        <div className={ styles.container }>
          <div className={ styles.row } style={{backgroundColor:"#53009e6e"}}>
            <div className={ styles.column }>
            {this.state.updateText.map((row, index) => (
                <Stack tokens={stackTokens}>
                  <table style={{width:"100%"}}>
                    <tr>
                      <th>Title</th>
                      <th>Supplier Name</th>
                      <th>Replacement For</th>
                    </tr>
                    <tr>
                      <td> <TextField underlined value={row.title} onChanged={(textval) => { row.title = textval }} style={{backgroundColor:"red"}} ></TextField> </td>
                      <td> <TextField underlined value={row.suppliername} onChanged={(textvalue) => { row.suppliername =  textvalue }} style={{backgroundColor:"yellow"}} ></TextField> </td>
                      <td> <TextField underlined value={row.replacementfor} onChanged={(replay)=>  { row.replacementfor= replay }} style={{backgroundColor:"blue"}}></TextField> </td>
                    </tr>
                  </table>
                  {/* title */}

                  {/* Supplier Name */}
                  <PrimaryButton text="Update" onClick={() => this._updateClicked(row)} />
                  <DefaultButton text="Delete" onClick={() => this._deleteClicked(row)} iconProps={DelIcon} />
                </Stack>
              ))}

                <br></br>
                <hr></hr>
                <label>Create New Item</label>
                <Stack horizontal tokens={stackTokens}>
                  <TextField label="Title" underlined value={this.state.addText} onChanged={(textval) => this.setState({ addText: textval })} ></TextField>
                  <PrimaryButton text="Save" onClick={this._addClicked} iconProps={AddIcon} />
                  <DefaultButton text="Clear" onClick={this._clearClicked} iconProps={ClearIcon} />
              </Stack>
            </div>
          </div>
        </div>
      </div>
    );
  }
     //  Event start
     async _getListItems() {
      const allItems: any[] = await sp.web.lists.getByTitle("vendor").items.getAll();
      console.log(allItems);
      let items: IListItem[] = [];
      allItems.forEach(element => {
        items.push({ id: element.Id, title: element.Title, suppliername: element.addSupplierName, replacementfor: element.replacementfor });
      });
      this.setState({ updateText: items });
 
    }

    @autobind
    async _updateClicked(row: IListItem) {
      const updatedItem = await sp.web.lists.getByTitle("vendor").items.getById(row.id).update({
        Title: row.title,
        suppliername: row.suppliername,
        replacementfor: row.replacementfor,
      });
  
    }
    @autobind
    async _deleteClicked(row: IListItem) {
      const deletedItem = await sp.web.lists.getByTitle("vendor").items.getById(row.id).recycle();
      this._getListItems();
    }
    
    @autobind
    async _addClicked() {
      const iar: IItemAddResult = await sp.web.lists.getByTitle("vendor").items.add({
        Title: this.state.addText
      });
      this.setState({ addText: '' });
      this._getListItems();
    }
    @autobind
    private _clearClicked(): void {
      this.setState({ addText: '' })
    }
}
