/*eslint-disable import/no-unresolved*/
import React from 'react'
import Match from 'react-router/Match'
import Link from 'react-router/Link'
import Redirect from 'react-router/Redirect'
import Router from 'react-router/BrowserRouter'
import Login from './Login'
import Auth from './Auth'

// -- not used; remove
const fakeAuth = {
  isAuthenticated: false,
  authenticate(cb) {
    this.isAuthenticated = true
    setTimeout(cb, 1000) // fake async
  },
  signout(cb) {
    this.isAuthenticated = false
        setTimeout(cb, 1000) // weird bug if async?
  }
}
// -- end

class AuthExample extends React.Component {

constructor(props) {
    super(props)
    this.state = {
        isSignedIn: false
    }
    this.updateState = this.updateState.bind(this)
    this.signOut = this.signOut.bind(this)
}

updateState() {
    this.setState({isSignedIn: !this.isSignedIn})
}

signOut() {
    this.setState({isSignedIn: false})
}

render() {
    return (
        <Router >
            <div>
                <h4> Excel REST Sample </h4>
                <Link to="/protected">Start </Link>
                {'|'}
                <Link to="/public"> Learn more </Link>

                <Match pattern="/public" component={Public}/>
                <Match pattern="/protected" component={Protected}/>
                <Match pattern="/auth" exactly={false} render={(props) => (
                        <Auth {...props} parentHook={this.updateState}/>) }
                />
            </div>
        </Router>)
  }

}
const NoMatch = ({ location }) => (
  <div>
    <h3>No match for <code>{location.pathname}</code></h3>
  </div>
)
const Protected = () =>  {
    return (<div> <Login/> </div>)
}
const Public = () => <a href='https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel'><h3>Explore Excel API on Graph.Microsoft.IO</h3></a>

export default AuthExample
