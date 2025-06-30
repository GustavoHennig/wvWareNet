# C# Porting Guidelines

## Primary Objective
Ensure that the translated C# code in the `WvWareNet` project accurately reflects the logic and functionality of the original `wvWare` C project.

## Comment Handling
- **Preservation**: 
    - Do not remove existing comments from the current C# code. They provide essential context.
    - Do not touch the original `wvWare` C project.
- **Enhancement**: You may update comments only to:
    - Improve clarity and readability.
    - Add new, relevant information (e.g., notes on C# specific implementation details).
    - Correct any inaccuracies in the original comments.

## General Instructions
- Follow standard C# coding conventions and best practices.
- Do not hardcode strings trying to simulate the expected result.
- Do not use reflection.
- The resulting code should be clean, readable, and maintainable.
