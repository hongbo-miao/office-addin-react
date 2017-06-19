import React from 'react';
import ReactDOM from 'react-dom';
import App from './App';
import registerServiceWorker from './registerServiceWorker';
import './index.css';

const Office = window.Office;

Office.initialize = () => {
  ReactDOM.render(<App />, document.getElementById('root'));
};

registerServiceWorker();
