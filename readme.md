Clone the github: git clone git@github.com:Ekugelsk/CS130-Wi24.git

(You may need to put your username instead of mine)

Next run these command:

`$ cd CS130-Wi24`

`$ python -m venv venv`

`$ echo -e $'venv\n.gitignore' >> .gitignore`

To activate the python venv: `source venv/bin/activate`

## Activate this everytime you begin to work on the project (you can setup VScode to automatically do this for you)

We will keep the requirments.txt up todate. This is how we will make sure we are all using the same packages.
Thus once you activate the environment you should run the following:

`$ pip install -r requirements.txt`

Make sure to update the requirements.txt everytime you install a new package, then push the file to GH.
To do this use the following command:

`pip freeze > requirements.txt`

### Branch Info

We are going to use a protected main branch. This means you can no longer push and pull to the main branch of github.\
Instead you must checkout a new branch when you want to begin working on something new.\

**Create and switch to a new branch** \
`git checkout -b <branch_name>`

**To switch between branches** \
`git checkout <branch_name>`

While working on this new feature you can push and pull to this new branch as much as possible (I'll get into pulling and keeping up with main in a second)

**To pull from your branch's remote repository**\
`git pull origin <branch_name>`

**To push to your branch's remote repository**\
`git push origin <branch_name>`

#### _If there are changes in main and you want to pull them into your branch_

First, switch to the main branch `git checkout main`\
Second, `git pull`\
Third, switch back to your branch `git checkout <branch_name>`\
Fourth, pull from main `git pull origin main`\
Last, push your changes with main update to remote `git push origin <branch_name>`

Once you've implemented that feature (we will keep this relatively short so PR are small), you need to make a pull request (PR).\
We do this so we can have multiple eyes looking over the code before we merge into main.\
Push your code to remote (make sure you're on your branch) `git push origin <branch_name>`\
Then on github you'll have to make the pull request. I can show you how once we get there. Don't forget to delete your branch once the PR was approved.
