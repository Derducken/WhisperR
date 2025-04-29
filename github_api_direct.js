const { Octokit } = require('@octokit/rest');

class GitHubAPI {
  constructor(token) {
    this.octokit = new Octokit({ auth: token });
  }

  async createRepo(name, description = '', isPrivate = false) {
    return this.octokit.repos.createForAuthenticatedUser({
      name,
      description,
      private: isPrivate
    });
  }

  async getRepos() {
    return this.octokit.repos.listForAuthenticatedUser();
  }

  async createFile(repo, path, content, message = 'Initial commit') {
    return this.octokit.repos.createOrUpdateFileContents({
      owner: repo.owner.login,
      repo: repo.name,
      path,
      message,
      content: Buffer.from(content).toString('base64')
    });
  }
}

module.exports = GitHubAPI;
