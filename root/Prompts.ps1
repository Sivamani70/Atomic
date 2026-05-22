function Get-Prompt {
    $Prompts = @(
        "Choose wisely! I haven't added any checks here, so if you give me a garbage name, I'll give you a garbage result.",
        "Name the child. Warning: Zero validation logic detected. If you use a colon, we both go down with the ship.",
        "Enter a filename. I'm too tired to sanitize your input, so try to be a responsible adult for once.",
        "File name, please! (Note: I've provided zero guardrails. Use a forbidden character and watch the world burn.)",
        "Give me a name. If you break the script with a weird character, I'm not helping you clean up the mess.",
        "The 'No-Validation' Zone: Enter a proper filename or prepare for an unplanned exit strategy.",
        "Your filename choice determines if we finish this shift in glory or in the debugger. No pressure.", 
        "I haven't did any checks. If this crashes, we're both telling the boss it was a hardware failure.",
        "It's too early for RegEx. Just give me a normal name and nobody gets hurt."
        
    )
    $ChosenPrompt = $Prompts | Get-Random
    return $ChosenPrompt
}