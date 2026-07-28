# rbPfPeek

<img alt="rbPfPeek" src="https://github.com/user-attachments/assets/cd7514f4-560b-49a6-bc2c-75651200fd0e" />

Please read the comments in the topmost part of the script to find detailed and updated information on configuration.

# Features
- The portfolio assets can be read from a CSV exported from Directa (italian broker), or manually inputed in the `$manualAssets` splatted variable, with quantity and average price.
- UI partially adapts to the console window size.
- Current market data is retrieved from the web.
- Runs on any Powershell version (tested in Windows Powershell v5.1 and Powershell v7.6.4).

# Why?
This exists because I love [Ticker](https://github.com/achannarasappa/ticker), but i wanted some more information on screen.

And also because I wanted to verify if all the hype on "vibe coding" meant something... 
Oh well actually it did! 
It all started playfully: I dropped a screenshot of Ticker in Claude, asking to build a Powershell script which could give me the same output, "let's see if something happens"... 2 minutes later I had the same output I have in Ticker!
So I decided to add stuff I missed from Ticker:
- the graphs
- the CSV parsing feature
- some cosmetic changes done manually
... and after 38 iterations/releases, all done locally in small free time spans, I decided to try GitHub to release it.

# Disclaimer
This was 98% built by the free version of these AI models: Claude, Gemini, Deepseek, and a bit of Qwen run locally.

So please expect errors.

Please report them if you can. I am not a developer, but I am willing to have my published stuff well working. 
Thank you!
