# 3-Diamond-Casino
The Excel Casino Game made on Mac

Welcome to the 3 Diamond Casino, the Mego Casino located at Applications://Microsoft Excel, just 2 exits down from Applications://VLC on I-95! The casino has multiple games and can ~~loose~~ win lots of money in. All games are made within excel and use no outside programs. There are currently 4 games that can be played and those are: Blackjack, Poker, Slots, and a Simulated Stock Market. And yes, the Bots in Poker are smart. A bit too smart. All is explained in detail down beleow. 

I did this project to try and see what Excel can do, how powerfull is it? Can is be used to create games? Well I guess I proved myself correct, right? Kinda. VBA sucks and I cant wait to be done with it. Excel is'nt made to do these things, and I had alot of settbacks when using it. Not only does it have a weak laungage, its also very slow. I own a reletivly new Mac and there are alot of times were Excel would freeze up on me and refuse to work. I've put alot of time into this project and learned alot, hopefully my wasted hours trying to figure out why a certian VBA function does'nt work just to find out no one mentioned that its not built into the Mac Excel library can give you some entertainment.

## Transfer Market:
Also known as the Home Page, this is the place to transfer money from diffrent games. Winning in games will update the Transfer Market and moving money in the Market will update the values on the Games sheet. The vault is a place just to store cash so you have some on hand. There is no loss in money when transfering between games so feel free to trade at will :)

Their are 3 main Macros on this page: `transferCash`, `updateValuesOut`, and `updateValuesIn`. 

`transferCash` gets the Cell Range of the button to find the value it will be subtracting and then asks the user where it wants to move it. VBA uses that Input to then    add the amount moved to this cell.

`updateValuesOut` is used to update the values of the "Total Money" cell in each game. This is called after money is trasnfered and when a game is started.

`updateValuesIn` is used to update the Transfer Market's values for each game. This is called after playing a game is finished.

## Blackjack
Nicknamed "Black Diamonds" on this side of town, the game of Blackjack has been locked down for a 4 year aggrement to stay at our Casino. The game plays just as normal Blackjack would, without Insurance and Splitting. Splitting? Why that is being added in a future installment of the game. 

To play a card game, we need a deck of cards. Not just any deck, but a random deck of cards. To do this, I first created an array for the deck of cards, each card having a value,suit and an ID attached to it. I then used the `RANDBETWEEN(1,208)` function to create a list of random values and then a `UNUIQE()` function to get the list of random ID's. Then use a `XLOOKUP()` to match the random ID to the card value and you have a random deck! But wait,`RANDBETWEEN()` is a volitile function. Having a deck of random cards updating every time the workbook is changed is'nt good. I bypass this trouble by using VBA. The `CreateBlackJackDeck()` function gets the value of this dynamic array and sets our static array equal to it. Now we have a randomly shuffled deck whenever we need one.

The game istself is straightfoward, `DealCardsBlackJack()` shows you your cards and one of the dealer's. Within the blackJackGame() sub (the one that calls `DealCardsBlackJack()`), the value of the dealer's other card is saved for later. Promts are then given to the player and depending on their awnser, Excel responds. If player is still in the game, then `dealerAction()` is called. The dealer hits until 17. Payout is calculated and money is only lost when you loose, there is no buy in.

## Poker

## Slots

## Stocks
