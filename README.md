# 3-Diamond-Casino
The Excel Casino Game made on Mac (both are'nt made to make games)

Welcome to the 3 Diamond Casino, the Mego Casino located at Applications://Microsoft Excel, just 2 exits down from Applications://VLC on I-95! The casino has multiple games and can ~~ lose ~~ win lots of money. All games are made within excel and use no outside programs. There are currently 4 games that can be played and those are Blackjack, Poker, Slots, and a Simulated Stock Market. And yes, the Bots in Poker are smart. A bit too smart. All is explained in detail down below. 

I did this project to try and see what Excel can do, how powerful is it? Can it be used to create games? Well, I guess I proved myself correct, right? Kinda. VBA sucks and I can't wait to be done with it. Excel isn't made to do these things, and I had a lot of setbacks when using it. Not only does it have a weak language, but it's also very slow. I own a relatively new Mac and there are a lot of times were Excel would freeze up on me and refuse to work. I've put a lot of time into this project and learned a lot, hopefully, my wasted hours trying to figure out why a certain VBA function doesn't work just to find out no one mentioned that it's not built into the Mac Excel library can give you some entertainment.

## Transfer Market:
Also known as the Home Page, this is the place to transfer money from different games. Winning in games will update the Transfer Market and moving money in the Market will update the values on the Games sheet. The vault is a place just to store cash so you have some on hand. There is no loss in money when transferring between games so feel free to trade at will :)

There are 3 main Macros on this page: `transferCash`, `updateValuesOut`, and `updateValuesIn`. 

`transferCash` gets the Cell Range of the button to find the value it will be subtracting and then asks the user where it wants to move it. VBA uses that Input to then add the amount moved to this cell.

`updateValuesOut` is used to update the values of the "Total Money" cell in each game. This is called after money is transferred and when a game is started.

`updateValuesIn` is used to update the Transfer Market's values for each game. This is called after playing a game is finished.

## Blackjack
Nicknamed "Black Diamonds" on this side of town, the game of Blackjack has been locked down for a 4-year agreement to stay at our Casino. The game plays just as normal Blackjack would, without Insurance and Splitting. Splitting? Why that is being added in a future installment of the game.

To play a card game, we need a deck of cards. Not just any deck, but a random deck of cards. To do this, I first created an array for the deck of cards, each card having a value, suit, and an ID attached to it. I then used the `RANDBETWEEN(1,208)` function to create a list of random values and then a `UNUIQE()` function to get the list of random IDs. Then use an `XLOOKUP()` to match the random ID to the card value and you have a random deck! But wait, `RANDBETWEEN()` is a volatile function. Having a deck of random cards updating every time the workbook is changed isn't good. I bypass this trouble by using VBA. The `CreateBlackJackDeck()` function gets the value of this dynamic array and sets our static array equal to it. Now we have a randomly shuffled deck whenever we need one.

The game itself is straightforward, `DealCardsBlackJack()` shows you your cards and one of the dealers. Within the blackJackGame() sub (the one that calls `DealCardsBlackJack()`), the value of the dealer's other card is saved for later. Prompts are then given to the player and depending on their answer, Excel responds. If the player is still in the game, then `dealerAction()` is called. The dealer hits until 17. Payout is calculated and money is only lost when you lose, there is no buy-in.

## Slots
The slot game is big, and we needed a new and original concept. So that's what we did. We named the Slot Machine "Devious Diamonds". Revolutionary stuff I know. The game is sure to attract fans all over the Excel sphere. I already got one friend addicted to gambling. "Andrei, I spent so much time playing other slot games but yours is the best, I can actually understand what's going on in yours. The other games are alight I guess"

The game is pretty simple, the simplest of them all. I first have an array filled with symbols and an ID attached to each. There are 36 symbols in total but symbols like cherries show up 8 times while the Diamond symbol only shows up once. In our dynamic array, we have each cell use a `RANDBETWEEN(1,36)` as the lookup value for our `XLOOKUP()` so we can get a random value and then display which symbol it corresponds to. We do the same thing in Blackjack, where we get the value of the dynamic array and set the static array equal to it.
<Finsih>
  
## Stocks
If you are more of an investment kinda guy, here you go. The Blue Market offers you a chance to invest in 15 different stocks, making large money. From a man who has made $500,000 out of $1,000, the goal is to go all in on `TYU`, its the best stock out there (i only own 150,000 shares and bought it all for $3.43). The best part is the data analysis (you know, the thing excel is made to do)! Of course, not much is there yet but there will be in update `Alpha v5`.
  
The stock market is also pretty simple with the biggest problem being the lag. Every time `stockRun` runs, it freezes up the entire workbook for a solid second or two. For each stock, there are 4 important values: `Stock Value`, `Trend`, `Trend Length`, and `Noise`.
  
`Trend` is calculated by using `NORM.INV(RAND(),0,5)` to get a random value on a curve. This number is recalculated once the `Trend Length` is equal to 0. 

The `Trend Length` is also a `NORM.INV` value but `ROUND()` ed and `ABS()` ed. The `Trend Length` is updated once it hits zero. 

`Noise` is the same as Trend just smaller, this is recalculated every time interval. 

 Each time interval, the `Trend` and `Noise` is added to the `Stock Value` and the `Trend Length` goes down by one. All of these values are unique to each stock and the new `Stock Value` is added to the respective data table. The graphs are connected to each of the data tables and might take time to update. Each stock has a buy and sell button. The Blue Highlighted cell tells you how much of that stock you have. Every time you buy and sell, the receipt table will update with info with the date, which Stock was affected, did you buy or sell, the amount bought/sold, the Value of the stock at the time, and the money changed (this will be negative for buying and positive for selling). Sadly this is my favorite part of the entire project. It even looks like a receipt :)
## Poker
