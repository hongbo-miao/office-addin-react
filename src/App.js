import React, { Component } from 'react';
import logo from './logo.svg';
import './App.css';

// const Excel = window.Excel;

class App extends Component {
  constructor(props) {
    super(props);

    this.state = { color: 'red' };

    this.changeColor = this.changeColor.bind(this);
    this.color = this.color.bind(this);
  }

  changeColor(event) {
    this.setState({ color: event.target.value });
  }

  color() {
    window.Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = this.state.color;
      await context.sync();
    });

    this.setState(prevState => ({
      color: this.state.color
    }));
  }

  render() {
    const colors = ['red', 'blue', 'yellow'];

    const listColors = colors.map(color =>
      <option key={color}>
        {color}
      </option>
    );

    return (
      <div className="App">
        <div className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <h2>Office.js ❤ React</h2>
        </div>

        <select value={this.state.color} onChange={this.changeColor}>
          {listColors}
        </select>

        <button onClick={this.color}>Color Me</button>
      </div>
    );
  }
}

export default App;
