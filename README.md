## Why
Hi! This project was created during my employment at Municipal Council. I was assigned to adjust balance of contrahents and add taxes to it from previous years due to new regulations in Poland. There were a couple of thousand codes to correct, which was very monotonous. Also, the software I had to use (called **KSAT2000i**) was prone to human errors, which wasn't very pleasant. So I decided to make it more fun, more precise and use this opportunity to learn something new.

## How
Unfortunetly, I'm not able to show off any previews of this software in action, except the GUI, so I'll try my best to describe everything using words.
- Enter contrahents code
- Check balance account and look for any arrears and what amount they relate to, including dates
- Search for contracts related to commune with type of receivable "G-PWŁ" and "G-WŁ"
- Generate document corrections for particular years in "G-WŁ"
- Create new footnote in "G-PWŁ" with adequate dates
- Create a financial document, ready to be accounted
- Send all invoices to database
- Check code in excel and prepare next one
- Repeat

Computer does about 90% of work for me. All I have to do is analyse contrahent's account if his situation is simple enough for script to handle and give exact positions of names, contracts, footnotes (if it's `n-th` counting from top) etc.

## Conclusion

&ensp; It was my first time using a scripting language and I'm happy with the end result. Code might be a little messy here and there, but it's something I'm looking to improve upon and educate myself more
&ensp; I've used **AutoHotkey** due to it's simplicity, ability to make a portable version,without a need of standard installation process and how straight forward documentation was, with many online topics. It's definitely not my last ride with this language.
&ensp; It was also an opportunity to use a new source code editor **Notepad++**. I didn't dig much into it, but customization capabilities were very nice and I'll most likely start using it instead of traditional OS notepads.

