# Debugging

When writing code, it is normal to make mistakes. Some mistakes are caused by simple typographical "syntax errors", while others may result from broader misunderstandings or oversights on our part about the logic of our application or the capabilities of the underlying technology we are working with. Others may even be caused by flaws in the underlying technology itself.

These mistakes in software, or differences between intended and actual functionality, are known as **bugs**. And the process of identifying and fixing them is called **debugging**.

## Debugging Strategies

Regardless of why an error is happening, the important thing to remember is to stay calm and stay positive. There's a good chance the error is fixable! And there's a good chance you aren't the only one who has ever seen it before.

  1. Your first objective when encountering an error should be to **communicate** what is happening, and how it differs from your expectation about what should happen. Ways to help you communicate the issue include: writing down your observations, vocalizing your observations to yourself, discussing your observations with someone else, or preparing to write a question to post on an online chatroom or help forum.
  2. Once you have communicated the problem, your second objective is to **identify** its source and its cause. Use various debugging tools and techniques (see below) to find out which line of code is causing the problem, and why.
  3. Finally, once you have identified the source of the problem, your objective is to **fix** it. Try one strategy at a time, using scientific method to determine whether or not that approach worked before trying another approach.

## Debugging Tools

Certain tools available in MS Excel and the VBE window can be helpful to us during the debugging process.

### Message Boxes

Once you learn about [Message Boxes](message-boxes.md), you can use them to display variable values at various stages in your program's execution.

### The Immediate Window

Similar to displaying variable values in a message box, you could also "print" variable values in a specific place called the Immediate Window.

Be aware of the official docs on the [Immediate Window](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/immediate-window), but actually spend time reading and following along with this unofficial [Using the Immediate Window](https://www.excelcampus.com/vba/vba-immediate-window-excel/) guide, which is better.

> FYI: sometimes code that doesn't actually cause an error during the normal flow of program execution will cause an error when you are evaluating it in the Immediate Window. For this reason, you might not be able to fully rely on the Immediate Window. That being said, it is still helpful for quick debugging in many cases.

### Trace Code Execution

Reference this document on [Trace Code Execution](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/trace-code-execution) for techniques on how to "step-through" your program line-by-line to better understand its order of execution.

### The Locals Window

Another powerful, but perhaps less commonly used, tool to see how your variables values change over time is
[The Locals Window](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/locals-window).

## Getting Help

After trying to debug on your own, if you are still unsuccessful, its normal to ask for help. Here are the sources of help you might try, in order of priority:

  1. Consult the professor's reference materials (this site).
  2. Consult the official Microsoft documentation (although sometimes it can be confusing).
  3. Ask the Internet, primarily using search engines like Google and programming forums like [Stack Overflow](https://stackoverflow.com/). Start your search terms with the words "MS Excel VBA" for more relevant hits.
  4. Ask someone you know, like a classmate, friend, mentor, or the professor.
  5. Take a break and revisit the problem later. Sometimes a walk or shower or good night's sleep will bring a fresh perspective.
