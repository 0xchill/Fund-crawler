// App.jsx
import React from "react";
import {Button,Col,Row} from "mdbreact";
import 'mdbreact/dist/css/mdb.css';
import Download from '@axetroy/react-download';
import moment from 'moment'

export default class App extends React.Component {
  constructor(props){
    super(props);
    this.state={
      empty:true
    };
    this.handleClick = this.handleClick.bind(this);
  }

  handleClick(){
    fetch('/receiver', {
      method: 'GET',
    })
    .then((response) => response.blob())
    .then(()=>this.setState(()=>({empty:false})));
  };

  render () {
    return (
      <Row className="d-flex justify-content-center">
        <Col className="align-self-middle" md="2">
          <Button
            onClick={this.handleClick}>
            Générer excel
          </Button>
          {!this.state.empty &&
            <Button href='/download'>Télécharger</Button>
          }

        </Col>
      </Row>
    );
  }
}
