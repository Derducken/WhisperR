const { createMCP } = require('@modelcontextprotocol/mcp');

// Replace 'your_github_pat' with the actual PAT provided by the user.
const mcp = createMCP({
  name: 'GitHub API MCP Server',
  description: 'An MCP server for interacting with GitHub using its API.',
  tools: [
    {
      name: 'create_repo',
      description: 'Create a new repository on GitHub.',
      command: `curl -X POST https://api.github.com/user/repos -u ${process.env.GITHUB_USER}:${process.env.GITHUB_PAT} -d '{"name":"${repo_name}"}'`,
      requires_approval: true,
    },
  ],
});

module.exports = mcp;
