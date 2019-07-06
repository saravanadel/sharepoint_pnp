
//https://github.com/s-KaiNet/node-sp-auth/wiki/SharePoint%20Online%20addin%20only%20authentication

// Client Id: ca59feec-202d-4560-ae17-514af0637411
// Client Secret: S3Qnl1+ganlRmFYHHn4xBtdZjF4tbmm7j1e3bKseOgA=
// Title: ClientApp
// App Domain: localhost
// Redirect URI: https://redirect.url

const { Web, IPnpNodeSettings } = require('@pnp/sp');
const { PnpNode } = require('sp-pnp-node');

let config = {
    // siteUrl - Optional if baseUrl is in pnp.setup or in case of `new Web(url)`
    siteUrl: 'https://karyatechnologiesindia.sharepoint.com/sites/EmployeePortal/',
    authOptions: {
        clientId: 'fa4fbd17-8bee-4192-814b-d670dd2787e1',
        clientSecret: 'VOZQEdiuguezp7iSrW2kKr5QDsInwH31VZG1XCuDO4I=',
        realm: 'c113852b-6df0-4882-9ac4-da9aa82b8714'
    }
};

new PnpNode(config).init().then(settings => {

    // Here goes PnP JS Core code >>>

    const web = new Web(settings.siteUrl);

    // Get all content types example
    web.lists.getByTitle('Results').items.get()
    // .items.add({
    //     Title: 'Karthik2'
    // })

        .then(data => {
            console.log(data);
        })
        .catch(console.log);


    // <<< Here goes PnP JS Core code

}).catch(console.log);