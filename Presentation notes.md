# Introduction - audience, scope and goals

This presentation is for my former biology lab mates. We study drug resistance in Candida, mostly Candida albicans (although my area of research was Candida glabrata). Most lab members are non-technical, and only some have any formal programming training, while some picked it up along the way. Folks generate a lot of data - growth curves, disk diffusion assays, spot assays, colony growth assays are routine. More fancy data includes whole-genome sequencing and proteomics. We have to analyze and visualize this data by ourselves, for presentations and for publication, and there is no 'one shoe fits all' approach - everyone has different needs and ideas. Access to technical people like myself is limited, and in any case not practical for day-to-day tasks. Thus, most if not all people in the lab work with web chatbots (like ChatGPT or Claude.ai) to generate code that analyzes the data in R or Python. This can create many problems and 'technical debt' if you're not careful, but mostly it's just slow to write and slow to manage.

In recent weeks I've been experimenting with Antigravity, and it has really blown my mind. First, for color commentary - all of my friends from industry (programmers, bioinformaticians, cyber specialists and data analysts) - none have written actual code manually in almost a year now. And now I know why - tools like AGY allow you to develop your project interactively, in human language, with very fast feedback. Gemini itself has put it quite well:

Chatbots (Local or Web UI)
• Disconnected from files: Cannot see your workspace, raw data files, or existing scripts.
• Copy-Paste Friction: You copy a prompt -> paste code to a file -> run in terminal -> copy error back to browser.
• No Execution Power: The chatbot generates code in a vacuum without testing whether it actually runs or throws errors.
• Mental Model: Like texting a friend for advice who can't see your computer screen.

Antigravity Development Environment
• Native Workspace Access: Directly reads your CSVs, image directories, config files, and scripts.
• Autonomous Execution: Runs scripts directly in your local conda environment and inspects output.
• Self-Correction Loop: If code throws a traceback or bad shape, the agent reads the error, modifies code, and re-runs. This looping ability is probably the strongest feature of this tool, making it goal-oriented, not just question-oriented.
• Mental Model: Like having a dedicated junior bioinformatician whom you can task and teach very quickly.

Antigravity is best thought of as not an agent, but a platform, or a development environment, which is local to your computer, but is powered by an AI (which can also be from Anthropic, although it's part of the Google Gemini ecosystem and probably works best with Gemini).

In this presentation, I'd like to show you how I used AGY, that is, what can be accomplished. I'll then showcase a 'best practices' instruction manual for beginners, and finally I'll throw around some suggestions for what you can do with in the immediate future.

# What did I do with AGY?

## Project 1: the paper reproducer

To learn AGY, I wrote a project that could take as input a genomics paper (as a PDF or a URL), and then generate a sub-project that attempts to reproduce the results in that paper. That requires quite a few steps that the AI has to do itself:
1. Extract and install the tools used (with their exact versions). Tools can be quite old, and sometimes they have to be installed from source, with changes in their code to run on modern OSes.
2. Download the raw data from the relevant repositories. If the downloads are too slow - figure out how to download faster.
3. Figure out how the tools are run, in which order.
4. Generate the figures from the results of the tools, and compare them to the paper's figures.
5. All this requires optimizations (if running on a personal computer, there are space and runtime considerations) and adaptations (personal computer vs. SLURM vs. a bagillion other runtime environments).

The project is still in development, but already I got preliminary figures, to prove that the pipeline itself is working. And I didn't write a single line of code! 

[Include in this part - the branching execution pathway the AI figured out, and one of the figures.]

## Project 2: the petite paper figure generator

Of course, the big question is, how can this help us in our day-to-day work? Right now, I'm writing my petite paper. I accumulated years of data, dispersed through my computer, and no fewer than 5 Jupyter notebooks, all bloated with a lot of experimental analyses and visualizations. The publishable figures are hiding there, somehwere, but it takes me a while even to remember where they are, let alone initialize the notebook properly. In addition, the color palettes between the figures were not coordinated, and making them so would require quite a lot of refactoring work. Finally, if the journal (or another researcher) asked for the code and data, that would've been very hard to provide in a way that they could actually make work.

Refactoring all of this by hand would've taken me weeks. So with what I learned from the paper-reproducer project, I decided to try it with AGY. I did help it - I marked all of the cells where the real productions figure were. But other than that, I basically told it, "make this work" (not really - the prompts are important).

I now have a single point of execution - one script that can generate all or some of the figures on demand. Figure generation is hooked up to an online Google Sheets that stores the configuration of the color palette, and can be edited by collaborators. This makes it easy to test color palettes. It also collected and organized all of the raw data files into a single folder, with clear documentation of everything. And, of course, I used it to improve a lot of small things in my figures - legends, labels, sizes, visual artifacts, etc. Now I'm working on typesetting the figures. And again - not a single line of code written by me.

[Showcase all of these things!]

Final fun fact - when I shared this project with Anastasia, who works on Windows, it didn't work (I'm running it on macOS). The failure was completely silent - no warning, no error, no crash report. I was completley flabbergasted. AGY solved it all by itself, realizing that this was a silent crash within matplotlib that could be fixed by setting an envrionmental variable. This would've taken me hours to solve at the very least - AGY did it in 20 minutes.

<!-- ## Project 3: this very presentation! -->

# What do you need to work with it?

You do need strong domain knowledge - you need to know what you want, be able to identify incorrect results, and be ready to go into detail and explain to AGY what it did wrong and what should be fixed. You also need to be very careful with the core analysis code - you really want to make sure that AGY explains in detail what it does, and documents this in the code as it does it. For example, in my proteomics work, it incorrectly normalized my data, probably because I didn't explain what the data actually was, and how it was processed before. Johannes caught this, and we fixed it, but without a domain expert this would've probably slipped through the cracks, altering the final results.

You can - and should - ask it to explain its reasoning, write tests, and run it on known datasets, so that it or you can compare the results to a 'ground truth'. Trusting it blindly is still not a good idea. The responsibility for the results is still yours!

Technical knowledge also goes a long way. Programming is still not obsolete - it can help if you can have at least an intuitive understand of what it's doing, so you can catch it early if it messes up. More important is the architechtural knowledge. For example, making sure that no names or values are hard-coded, but are either computed or stored in textual configuration files for you to edit as required. Or making sure that there is no code duplication. Even if the AI writes a working project, it may contain a lot of bloat and poor architecture, which, in the long run, will make it burn tokens, run slowly, and give poor results as the project grows. But even here, stronger models can be used for occasional refactors.

You still want to work with text (code is also text!). You DON'T want the AI to edit images directly. Instead, give it the raw data, have it write code to analyze and visualize that data, and if you need any changes to the figure, have it edit the code that creates that figure. That's deterministic and easy to manage. AIs are not (yet) total magicians.

# An overview of some useful technical features and best practices

## Background on how AGY works

First, you create a project, that is associated with a folder. Everything within that folder is considered local, and thus safe to run. AGY will ask you if you accesses files outside that folder, so it may be prudent to copy what you need into the project folder. Then, you start a conversation - a persistent session, this works similiarly to how chatbots conversations work. You can also chose the model and 'thinking effort' per conversation. A few important points regarding conversations:

1. Each conversation starts completely fresh, without any memory. It can scan the folder and files within, but that takes tokens. However, you can 'seed' it with instructions via AGENTS.md (we'll talk about it soon), and you can point to important files and even other conversations with the '@' command.
2. Each conversation carries a context - everything that was said in the conversation prior. The context is finite, and once it fills up, the conversations are 'compacted' (summarized) and the previous context is discarded, so that the conversation can go on. Try not to get to that point, as performance degrades. Having many short conversations (each conversation about a particular feature, panel, or bug fix) is better than having one long one. Remember: for LLMs, only the context exists. Whatever isn't in it, doesn't exist. [Have a context visualizer here.]
3. Better models are more expensive (that is, they burn tokens faster). You can make AGY switch between models, so you can, for example, use a cheaper model for simple tasks like finding files, and a better model for complex tasks like debugging. You can also switch between thinking levels - from 'deep thinking' (slower, more expensive) to 'shallow thinking' (faster, cheaper).
4. You can even ask Gemini (or another chatbot) to engineer the prompt and choose a model. It can also give you tips, for example, which slash commands to use (see below).
5. While running, you can see what the agent is doing - what tools are running, and its thinking process - and you can stop it and correct course if needed.

## Managing the context

1. `AGENTS.md` - this is the key document that gets auto-loaded into every conversation of the project. Here you explain to the AI anything that it may require to work quickly and efficiently with your project - the project goal, overview of the files, which tools it can (and cannot!) run, rules and invariants (e.g.,  writing or visual style).
2. Mentions with the `@` command - you can point it to files, rules, and even prior conversations.
3. There's the `/compact` command, which you can call to compact the context when (a) it's getting filled up, and (b) when the last chunk of work finished properly, and there's no danger of compacting away something important. Note on tracking usage: context usage stats exist in the CLI environment, but I couldn't find it in the App, even though Gemini insists it's there and insists that knowing when to use `/compact` is important. Thus, I would keep an eye on this command and context usage in general.
4. Many short conversations are OK!

## Useful slash commands

- `/plan` - when starting a big chunk of work (e.g., analyzing a new file format), you can ask AGY to make a plan, which you can then read and edit iteratively, if needed. This gives the agent more precise instructions to follow, and you the peace of mind that it's on the right track.
- `/grill-me` - this is useful alongside `/plan`, but can be used without it. This instructs the agent to ask you about points it's unsure about before execution.
- `/learn` - once the agent has completed a task, this instructs it to learn from what it did, and update `AGENTS.md` and other relevant files, perhaps even creating a skill. Can be used without any instructions!

## Skills

Skills are like functions in programming, but for AI. They describe a specific task for the AI - how it should go about it, any guardrails it should, and the tools it should use. I personally only wrote one skill so far, for working with the figure color palette - it can generate a 'master' palette from very general instructions (e.g., from a website that lists the colors), and then populate the full configuration list. You can also import skills from others, and AGY comes with many built-in skills you can use.

## Git and GitHub

While not strictly necessary, working with git is really useful for keeping track of changes and rolling back if needed. It's also very useful for collaboration, as you can store your repository in GitHub (and make it private until you publish it). Make sure to not store any secrets (like passwords and access keys) in the repository, though!

Example: I shared my repository with Anastasia, who can now generate the figures, fix bugs or make any additions autonomously, and the code exchange between us is safe and controlled.

One advanced feature of working with git are worktrees, which allow you to run tasks in parallel, without interfering with other. It requires a small amount of architecture and thought, but is worth it for mid-sized projects and up.

This presentation is too short to discuss git, but I can give a separate presentation if you want.

## Safey and security

- The AI will ask you if it tries to access files outside its local folder. This can even become annoying. Still, you must be careful with running tools - as it CAN modify data on your computer, and modify or delete important files by mistake. This is where tools like git and Dropbox come in handy, as they keep archives. My recommendation is to copy your data files into your local folder, so the local copy is accessible to the AI, and you have an archival copy elsewhere (on Dropbox or the NAS or whatever).
- The AI can hide access to outside the folder inside code it generates. It tries not to, but this is one more reason to review what it writes before running it.

# Next steps

- Right now, the lab has a few tools, like OD Normalizer, DDA panel creator, and older and bigger ones, like Scanalyzer and even diskImageR, that nobody really maintains. These can now be maintained and improved by AGY, although it requires a certain amount of setup. My homework for Ze'ev - should he accept it - is to feed the GitHub issues of the DDA panel creator to AGY - I'm sure you can finisih all of them within a week.
- Shani - I see you working with Claude's UI! I recommend playing around with AGY and seeing if you can consolidate your analysis work into a single project, and see how it can benefit you to writing better analyses and visualizations.
- Sofya - I'm positive AGY can be used for better exploratory analysis of transposon libraries. We still need quality assessments on those libraries, and I don't think we ended up with something final.
- Alona - I know you already have a Claude subscription; see if you can plug Anthropic's Claude models into AGY (or test Anthropic's CLI agent Claude Code).

- AGY itself is free, and the Gemini models powering it do have a free tier , but free quotas can run out quickly with heavy use. TAU provides a 550 NIS per year Gemini Pro subscription (though verify whether TAU's license covers developer API quotas or just the Google Workspace web chat interface). By comparison, standard private developer subscriptions or API billing runs around ~75 NIS / month (~900 NIS / year). TAU also offers academic access to OpenAI and Microsoft models (though Microsoft's Copilot is considered inferior for code). You can also plug in Anthropic's Claude models (e.g., Claude 3.5/3.7 Sonnet or Opus) into AGY via API keys. While many consider Anthropic's coding models superior for certain tasks, for my biology pipelines Gemini Pro was completely sufficient and very cost-effective.