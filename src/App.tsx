import React, { Component } from "react";
import { BrowserRouter as Router, Route, Link } from "react-router-dom";
import { Silent } from "./components/silent";
import { SilentStart } from "./components/silent-start";
import { SilentEnd } from "./components/silent-end";

class App extends Component {
	render() {
		return (
			<Router>
				<div>
					<Route exact path="/" component={Silent} />
					<Route path="/start" component={SilentStart} />
					<Route path="/end" component={SilentEnd} />
				</div>
			</Router>
		);
	}
}

export default App;
