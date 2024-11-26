import React from 'react';
import './App.css';
import ExcelTemplate from './components/ExcelTemplate';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>React Excel Template App</h1>
      </header>
      <main>
        <ExcelTemplate />
      </main>
    </div>
  );
}

export default App;
