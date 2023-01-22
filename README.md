# 3-Diamond-Casino
The Excel Casino Game made on Mac (both aren't made to make games)

Welcome to the 3 Diamond Casino, the Mego Casino located at Applications://Microsoft Excel, just 2 exits down from Applications://VLC on I-95! The casino has multiple games and can ~~lose~~ win lots of money. All games are made within excel and use no outside programs. There are currently 4 games that can be played and those are Blackjack, Poker, Slots, and a Simulated Stock Market. And yes, the Bots in Poker are smart. A bit too smart. All is explained in detail down below. 

I did this project to try and see what Excel can do, how powerful is it? Can it be used to create games? Well, I guess I proved myself correct, right? Kinda. VBA sucks and I can't wait to be done with it. Excel isn't made to do these things, and I had a lot of setbacks when using it. Not only does it have a weak language, but it's also very slow. I own a relatively new Mac and there are a lot of times were Excel would freeze up on me and refuse to work. I've put a lot of time into this project and learned a lot, hopefully, my wasted hours trying to figure out why a certain VBA function doesn't work just to find out no one mentioned that it's not built into the Mac Excel library can give you some entertainment.



## Transfer Market:
Also known as the Home Page, this is the place to transfer money from different games. Winning in games will update the Transfer Market and moving money in the Market will update the values on the Games sheet. The vault is a place just to store cash so you have some on hand. There is no loss in money when transferring between games so feel free to trade at will :)

There are 3 main Macros on this page: `transferCash`, `updateValuesOut`, and `updateValuesIn`. 

`transferCash` gets the Cell Range of the button to find the value it will be subtracting and then asks the user where it wants to move it. VBA uses that Input to then add the amount moved to this cell.

`updateValuesOut` is used to update the values of the "Total Money" cell in each game. This is called after money is transferred and when a game is started.

`updateValuesIn` is used to update the Transfer Market's values for each game. This is called after playing a game is finished.

## Blackjack:
Nicknamed "Black Diamonds" on this side of town, the game of Blackjack has been locked down for a 4-year agreement to stay at our Casino. The game plays just as normal Blackjack would, without Insurance and Splitting. Splitting? Why that is being added in a future installment of the game.

To play a card game, we need a deck of cards. Not just any deck, but a random deck of cards. To do this, I first created an array for the deck of cards, each card having a value, suit, and an ID attached to it. I then used the `RANDBETWEEN(1,208)` function to create a list of random values and then a `UNUIQE()` function to get the list of random IDs. Then use an `XLOOKUP()` to match the random ID to the card value and you have a random deck! But wait, `RANDBETWEEN()` is a volatile function! Having a deck of random cards updating every time the workbook is changed isn't any good. I bypass this trouble by using VBA. The `CreateBlackJackDeck()` function gets the value of this dynamic array and sets our new static array equal to it. Now we have a randomly shuffled deck whenever we need one :thumbsup:

The game itself is straightforward, `DealCardsBlackJack()` shows you your cards and one of the dealers. Within the blackJackGame() sub (the one that calls `DealCardsBlackJack()`), the value of the dealer's other card is saved for later. Prompts are then given to the player and depending on their answer, Excel responds. If the player is still in the game, then `dealerAction()` is called. The dealer hits until 17. Payout is calculated and money is only lost when you lose, money is'nt taken once the game starts, like in the real game (I dont know, I only played by myself).

## Slots:
The slot game is big, and we needed a new and original concept. So that's what we did. We named the Slot Machine "Devious Diamonds"! Never seen before in the compettitive Excel casino market. The game is sure to attract fans all over the Excel sphere. I can already hear the "This workbook contains macros. Do you want to disable macros before opening this file?" pop up on screens across the world. I already got one friend addicted to gambling. Here is his quote: "Andrei, I spent so much time playing other slot games but yours is the best, I can actually understand what's going on in yours. The other games are alight I guess."

The game is pretty simple, the simplest of them all. I first have an array filled with symbols and an ID attached to each. There are 36 symbols in total with varing amount of each sybmol in trhoughout. In our dynamic array, we have each cell use a `RANDBETWEEN(1,36)` as the lookup value for our `XLOOKUP()` so we can get a random value and then display which symbol it corresponds to. We do the same thing in Blackjack, where we get the value of the volitile array and set the static array equal to it.

Now look, I know very little about probabilties ok. I want to. I would love to, but I dont. The calculation for payouts is wrong ok. Ive spent alot of time talking to chatGPT, talking to smarter pepole, and doing trial and error. *Currently* this is how payouts are calculated. The winning cases are getting 3,4, or 5 of a single symbol in any row (its a 4x5 grid). The chances of each winning case are calculated:

sP^aM * (aM Choose 5) *4

sp = Symbol Perentage

am = Amount of Times the Symbol showed up

## Stocks:
If you are more of an investment kinda guy, here you go. The Blue Market offers you a chance to invest in 15 different stocks, making large money. From a man who has made $500,000 out of $1,000, the goal is to go all in on `TYU`, its the best stock out there (i only own 150,000 shares and bought it all for $3.43). The best part is the data analysis (you know, the thing excel is made to do)! Of course, not much is there yet but there will be in update `Alpha v5`.
  
The stock market is also pretty simple with the biggest problem being the lag. Every time `stockRun` runs, it freezes up the entire workbook for a solid second or two. For each stock, there are 4 important values: `Stock Value`, `Trend`, `Trend Length`, and `Noise`.
  
`Trend` is calculated by using `NORM.INV(RAND(),0,5)` to get a random value on a curve. This number is recalculated once the `Trend Length` is equal to 0. 

The `Trend Length` is also a `NORM.INV` value but `ROUND()` ed and `ABS()` ed. The `Trend Length` is updated once it hits zero. 

`Noise` is the same as Trend just smaller, this is recalculated every time interval. 

 Each time interval, the `Trend` and `Noise` is added to the `Stock Value` and the `Trend Length` goes down by one. All of these values are unique to each stock and the new `Stock Value` is added to the respective data table. The graphs are connected to each of the data tables and might take time to update. Each stock has a buy and sell button. The Blue Highlighted cell tells you how much of that stock you have. Every time you buy and sell, the receipt table will update with info with the date, which Stock was affected, did you buy or sell, the amount bought/sold, the Value of the stock at the time, and the money changed (this will be negative for buying and positive for selling). Sadly this is my favorite part of the entire project. It even looks like a receipt :)

## Poker:
Now for poker, the biggest game of all. Not because it's complex, but because the bots need to actually do things, smart things. Who knew that creating smart things is hard?
  
Let's skip the boring part, the deck is made, players get cards, the player can Call/Raise/Fold, bots make an action, cards are drawn, winners are made, friendships are ruined, blah blah blah. The cool part is the smart bots. For the initial bets, a basic strategy exists so I use that to determine the bot's initial play. Each 2-Card hand has a strength of 1-4, 1 Being raised and 4 being Fold. A simple sort for each hand and then an XLOOKUP() gives each bot a move. After the first card is drawn, the fun begins. 

The main idea was to make the bot use the cards it has to then predict the chances it has at winning. Now, how? `Sheet1`, that's how.

`Sheet1` holds all the power in this workbook and is my baby of great ideas. Let me go through a quick history of `Sheet1`, starting with `Book3`, the mommy of `Sheet1`. `Book3` was a workbook that was being used at school on my Chromebook because the main file was too big to open on onedrive. My goal was to first create all possible 2-card combinations of a 104 poker card deck. Now if I were a bit smarter, I would wait until I got home to use Power Query (which I'm going to start learning), but I instead used a big ass formula (and since this is onedrive on a $2 Chromebook I was able to grill some chicken on the keyboard while I faced a computer that won't work for another 20 minutes). What I needed next was a way to get a hand type from 5 cards. 
I used `COUNTIF()`'s to handle it and made it work. I first counted how many of each value type in a hand (how many 2,3,4... cards in the hand) then the amount of each suit and used that information to get a hand type. Next, we get the hand type and translate that into a number (1-9, 1 being high card and 9 being Royal Flush). Now we can do what we want to do, and that creates all possible outcomes with the hand we are given and get the average hand strength. It's a crazy idea, but it might just work. We now have to add the 3 cards we know the all combinations of 2-Cards. This means 10,816 rows to get all possible hands. For each row, we have the hand identifier with the strength attached to it. Now we get the average strength and there we go! If it's above a 2 (better than a pair) the Bot stays in, if it's better than a 3 (2 Pair) then the bot raises. And that's the story of `Book3`, the mother of `Sheet1`. 
  
 
Funnily, poker AI was the last thing I did before I submitted this for the science fair. It was the night it was due and I needed the AI to work, remembering `Book3` being finished a couple of months before, I went right to work on implementing it into my code. One big change was that I was going to use the same sheet to calculate the strength of a hand when knowing 3 and 4 cards. This was an easy fix but only having 30 minutes to finish this huge project, insane music playing, and two scoops of Gfuel in my body makes my thinking stall if I dont know what to do. The fix was simple, checking to see if there was a 4th card, and then changing the cell reference accordingly. 
 
It's no joke that this is slow, there are 367,744 cells having to be calculated. I'm currently trying to optimize it to the max but it's not going well. My goal was to use another workbook to create a prepopulated table of the hand values for all possible hand types. This would work by going through all 3 card combinations and plugging them into the "known cards" section of the calculation, getting the strength, and putting the ID and strength into a table. With this table, I could use an `XLOOKUP()` instead of simulating all possible hand types.
The only problem is that it takes 2 seconds per hand and there are 140,608 hands. I dont have 78 hours straight to let my personal mac run this program. I'm pretty sure the CPU would melt. This was optimized as much as I could and doing a bunch of tricks to reduce calculation times (it was 3.5 seconds before).
So I guess I gotta learn a real language and make it there, SMH.
