// App.jsx
import React from "react";
import {Button,Col,Row} from "mdbreact";
import 'mdbreact/dist/css/mdb.css';

export default class App extends React.Component {
  constructor(props){
    super(props);
    this.handleClick = this.handleClick.bind(this);
  }

  handleClick(){
    var cars = [
      { "make":"Porsche", "model":"911S" },
      { "make":"Mercedes-Benz", "model":"220SE" },
      { "make":"Jaguar","model": "Mark VII" }
    ];
    fetch('/receiver', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(cars)
    })
    .then((response) => response.text())
    .then((data)=>{console.log(data)});
  };
  render () {
    return (
      <Row className="d-flex justify-content-center">
        <Col className="align-self-middle" md="2">
          <Button
            onClick={this.handleClick}>
            Click Me
          </Button>
        </Col>
      </Row>
    );
  }
}
