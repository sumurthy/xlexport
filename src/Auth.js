import React from 'react'
import './App.css';
import Login from './Login'

class Auth extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
                    isSignedIn: false,
                    authCode: "",
                    message: ""
                }
        this.updateRange = this.updateRange.bind(this)
        this.createTable = this.createTable.bind(this)
        this.applyFilter = this.applyFilter.bind(this)
        this.logout = this.logout.bind(this)
    }

    logout() {
        this.setState({isSignedIn: false})
    }

    async componentDidMount() {

        let params = this._extractParams(this.props.location.hash)
        if (params['access_token']) {
            //this.setState({access_token: "Bearer " + params['access_token']}, async () => {
            let response = await this._getDrive(params['access_token'])
            if (response.ok) {
                let body = await response.json()
                this.setState({ isSignedIn: true, authCode: "Bearer " + params['access_token'] })
                this.props.parentHook()
            }
        }
    }

    render () {
        if (!this.state.isSignedIn) {
            return <Login/>
        }
        else {
            return (<div>
                    <br/>
                    <button onClick={this.updateRange} className="Button">Add sales data </button>
                    <br/><br/>
                    <button onClick={this.createTable} className="Button">Create sales table</button>
                    <br/><br/>
                    <button onClick={this.applyFilter} className="Button">Apply top-5 filter</button>
                    <p>{this.state.message}</p>
                    <p><button onClick={this.logout}>Logout</button> </p>
                    </div>
                )
        }

    }

    // A Function to update the range values
    async updateRange() {
        // Oauth access token obtained using implicit auth flow
        // check out Graph.microsoft.io for details.
        let at = this.state.authCode
        var headers = new Headers()

        // Set request headers.
        headers.append('Accept', 'application/json')
        headers.append('Content-type', 'application/json')
        headers.append('Authorization', at)

        // Prepare request body
        let range = {
            values: [['Product ID', 'Product Name', 'Price Date', 'Retail Price Per Unit', 'Bulk Price Per Unit*', 'Units Sold (retail)', 'Units Sold (bulk)', 'TotalSales', 'Total Sales ($)'],
                        ['5', 'Shorts', '1/1/2012', '20', '$20 ', '629', '1,254', '1,883', '$37,660'], 
                        ['1', 'Shirts', '1/1/2012', '88', '$54 ', '734', '1,427', '2,161', '$141,650'], 
                        ['2', 'Sandals', '1/1/2012', '70', '$44 ', '744', '1,043', '1,787', '$97,972'], 
                        ['3', 'Umbrellas', '1/1/2012', '63', '$44 ', '681', '1,523', '2,204', '$109,915'], 
                        ['4', 'Water bottles', '1/1/2012', '35', '$27 ', '602', '1,822', '2,424', '$70,264'], 
                        ['1', 'Shirts', '2/1/2012', '55', '$44 ', '678', '1,515', '2,193', '$103,950'], 
                        ['2', 'Sandals', '2/1/2012', '83', '$54 ', '753', '1,005', '1,758', '$116,769'], 
                        ['3', 'Umbrellas', '2/1/2012', '34', '$34 ', '986', '1,069', '2,055', '$69,870'], 
                        ['4', 'Water bottles', '2/1/2012', '35', '$25 ', '848', '1,211', '2,059', '$59,955'], 
                        ['5', 'Shorts', '2/1/2012', '41', '$38 ', '980', '1,330', '2,310', '$90,720'], 
                        ['1', 'Shirts', '2/29/2012', '27', '$18 ', '533', '1,936', '2,469', '$49,239'], 
                        ['2', 'Sandals', '2/29/2012', '38', '$28 ', '952', '1,512', '2,464', '$78,512'], 
                        ['3', 'Umbrellas', '2/29/2012', '92', '$92 ', '956', '1,266', '2,222', '$204,424'], 
                        ['4', 'Water bottles', '2/29/2012', '43', '$36 ', '952', '1,390', '2,342', '$90,976'], 
                        ['5', 'Shorts', '2/29/2012', '98', '$73 ', '530', '1,452', '1,982', '$157,936'], 
                        ['1', 'Shirts', '3/31/2012', '38', '$28 ', '973', '1,415', '2,388', '$76,594'], 
                        ['2', 'Sandals', '3/31/2012', '50', '$36 ', '672', '1,105', '1,777', '$73,380'], 
                        ['3', 'Umbrellas', '3/31/2012', '24', '$23 ', '769', '1,629', '2,398', '$55,923'], 
                        ['4', 'Water bottles', '3/31/2012', '72', '$57 ', '985', '1,848', '2,833', '$176,256'], 
                        ['5', 'Shorts', '3/31/2012', '85', '$43 ', '721', '1,426', '2,147', '$122,603'], 
                        ['1', 'Shirts', '4/30/2012', '91', '$65 ', '603', '1,226', '1,829', '$134,563'], 
                        ['2', 'Sandals', '4/30/2012', '91', '$55 ', '892', '1,823', '2,715', '$181,437'], 
                        ['3', 'Umbrellas', '4/30/2012', '42', '$42 ', '611', '1,181', '1,792', '$75,264'], 
                        ['4', 'Water bottles', '4/30/2012', '85', '$43 ', '530', '1,039', '1,569', '$89,727'], 
                        ['5', 'Shorts', '4/30/2012', '82', '$71 ', '716', '1,249', '1,965', '$147,391'], 
                        ['1', 'Shirts', '5/14/2012', '34', '$31 ', '850', '1,548', '2,398', '$76,888'], 
                        ['2', 'Sandals', '5/14/2012', '64', '$40 ', '876', '1,663', '2,539', '$122,584'], 
                        ['3', 'Umbrellas', '5/14/2012', '33', '$30 ', '881', '1,149', '2,030', '$63,543'], 
                        ['4', 'Water bottles', '5/14/2012', '29', '$27 ', '802', '1,548', '2,350', '$65,054'], 
                        ['5', 'Shorts', '12/11/2013', '24', '$15 ', '824', '1,994', '2,818', '$49,686']
                    ]
        }
        let response = null

        // Make the API call (POST)
        try {
            // URL: POST ../workbook/worksheets('{name}')/tables/add
            response = await fetch("https://graph.microsoft.com/v1.0/me/drive/root:/book3.xlsx:/workbook/worksheets('test')/range(address='test!A1:I31')",
                {
                    headers: headers,
                    method: 'PATCH',
                    body: JSON.stringify(range)
                }
            )
            if (response.ok) {
                this.setState({message: 'Add sales data'})
            }

        } catch (e) {
            // Handle error
            console.log("Error during table creation");
        } finally {
            return
        }
    }

    // A Function to create table on an existing range using "fetch" REST client
    async createTable() {
        // Oauth access token obtained using implicit auth flow
        // check out Graph.microsoft.io for details.
        let at = this.state.authCode
        var headers = new Headers()

        // Set request headers.
        headers.append('Accept', 'application/json')
        headers.append('Content-type', 'application/json')
        headers.append('Authorization', at)

        // Prepare request body
        let tbl = {
            address: 'sales1!A1:I31',
            hasHeaders: true,
        }
        let response = null

        // Make the API call (POST)
        try {
            // URL: POST ../workbook/worksheets('{name}')/tables/add
            response = await fetch("https://graph.microsoft.com/v1.0/me/drive/root:/book3.xlsx:/workbook/worksheets('sales1')/tables/add",
                {
                    headers: headers,
                    method: 'POST',
                    body: JSON.stringify(tbl)
                }
            )
            if (response.ok) {
                this.setState({message: 'Sales table created'})
            }

        } catch (e) {
            // Handle error
            console.log("Error during table creation");
        } finally {
            return
        }
    }

    // A Function to apply filter on an existing table using "fetch" REST client
    async applyFilter() {
        // Oauth access token obtained using implicit auth flow
        // check out Graph.microsoft.io for details.
        let at = this.state.authCode
        var headers = new Headers()

        // Set request headers
        headers.append('Accept', 'application/json')
        headers.append('Content-type', 'application/json')
        headers.append('Authorization', at)

        // Prepare request body
        let tbl = {
            count: 5
        }
        let response = null

        // Make the API call (POST)
        try {
            // URL: POST ../tables('sales')/columns('TotalSales')/filter/applyTopItemsFilter
            response = await fetch("https://graph.microsoft.com/v1.0/me/drive/root:/book3.xlsx:/workbook/worksheets('sales')/tables('sales')/columns('TotalSales')/filter/applyTopItemsFilter",
                {
                    headers: headers,
                    method: 'POST',
                       body: JSON.stringify(tbl)
                }
            )
            if (response.ok) {
                // Handle error
                this.setState({message: 'Top 5 total sales filter applied'})
            }

        } catch (e) {
            console.log("Error during table creation");
        } finally {
            return
        }
    }

/* 
    This method parses out the auth code if one exists. \

*/

    _extractParams = (segment = '') => {
        let parts = segment.split('#');
        if (parts.length <= 0) return {};
        segment = parts[1]
        let params = {}
        let regex = /([^&=]+)=([^&]*)/g
        let matches = ''

        while ((matches = regex.exec(segment)) !== null) {
            params[decodeURIComponent(matches[1])] = decodeURIComponent(matches[2])
        }

        return params
    }

    async _getDrive(at="") {
        var headers = new Headers()
        headers.append('Accept', 'application/json')
        console.log("the access token is2: " + at)
        headers.append('Authorization', at)
        let response = null
        try {
            return await fetch('https://graph.microsoft.com/v1.0/me/drive', {
                headers: headers
            })
        } catch (e) {
            console.log("Error");
        } finally {
        }
    }
}

export default Auth;
