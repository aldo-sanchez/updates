const fetch = require('node-fetch');
const axios = require('axios')
let dotenv = require('dotenv')
const officegen = require('officegen')
const { generateExampleSlides } = require('./pptx');
const fs = require('fs')

dotenv.config()

const accessToken = process.env.git;
let test = `
  {
    viewer {
      repositories(affiliations: [COLLABORATOR], last:50) {
        edges {
          node {
            nameWithOwner
            refs(refPrefix: "refs/heads/", last:40) {
              edges {
                node {
                  name
                  target {
                    ... on Commit {
                      history(since: "2018-02-20T05:00:00.000Z",first: 100, author: {emails: "josealdosanchez@gmail.com"}) {
                        edges {
                          node {
                            id
                            authoredDate
                            message
                            author {
                              date
                              name
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  
  `
 
fetch('https://api.github.com/graphql', {
  method: 'POST',
  body: JSON.stringify({query: test}),
  headers: {
    'Authorization': `Bearer ${accessToken}`,
  },
}).then(res => res.json())
  .then(body => {
    // console.log(body)
    let total = []
    let array = body.data.viewer.repositories.edges
    array.forEach(a => {
      let names = a.node.nameWithOwner.split('/')
      let owner = names[0]
      let repo = names[1]
      if(owner === 'QRIGroup') {
        console.log(repo)
        let branches = a.node.refs.edges
        branches.forEach(b => {
          let histories = b.node.target.history.edges
          let branch = b.node.name
          if(histories.length > 0) {
            console.log(branch)
            total.push(...histories)
          }
        })
      }
    })
    let uniq = {}
    let uniqueCommits = total.filter(commit => {
      let { id } = commit.node
      if(uniq[id]) {
        return false
      }
      return uniq[id] = true
    })
    console.log('filtered', uniqueCommits)
    console.log(total.length, uniqueCommits.length)
    generateExampleSlides(uniqueCommits, (pptx) => {
      var out = fs.createWriteStream('./tmp/out.pptx');
      out.on('error', function (err) {
        console.log(err);
      });
      pptx.generate(out);
    })
  })
  .catch(error => console.error(error));