using System;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Tutorial
{
    /*
        TutorialService provides interactive guided learning for new users.
        
        This service implements the Teaching Mode feature (Feature C from IPD plan).
        It walks users through available voice commands step by step, providing
        contextual tips and celebrating successful command execution.
        
        Key features:
        - Guided tutorial with progressive steps
        - "How do I..." query handler for on-demand help
        - Contextual tips based on current application
        - Progress tracking and encouragement
        
        This feature is essential for accessibility users who are computer beginners
        and need patient, step-by-step guidance to learn voice control.
    */
    public class TutorialService
    {
        private readonly ILogger<TutorialService>? _logger;
        private TutorialState _state;
        private int _currentStepIndex;
        private readonly List<TutorialStep> _tutorialSteps;

        public event EventHandler<TutorialEventArgs>? TutorialEvent;
        public event EventHandler<string>? SpeechRequested;

        public bool IsActive => _state == TutorialState.Active;
        public int CurrentStepIndex => _currentStepIndex;
        public int TotalSteps => _tutorialSteps.Count;

        public TutorialService(ILogger<TutorialService>? logger = null)
        {
            _logger = logger;
            _state = TutorialState.Inactive;
            _currentStepIndex = 0;
            _tutorialSteps = InitializeTutorialSteps();
        }

        private List<TutorialStep> InitializeTutorialSteps()
        {
            // Define the tutorial progression
            // Each step teaches one concept and waits for the user to try it
            return new List<TutorialStep>
            {
                new TutorialStep
                {
                    Title = "Welcome",
                    Instruction = "Welcome to AI Companion! I'll teach you how to control your computer with your voice. " +
                                  "Let's start with the basics. Try saying 'Hello' to greet me.",
                    ExpectedCommands = new[] { "hello", "hi", "hey", "привет" },
                    SuccessMessage = "Great job! You've learned how to talk to me. Let's move on!",
                    Hint = "Just say 'Hello' or 'Hi' to continue."
                },
                new TutorialStep
                {
                    Title = "Opening Applications",
                    Instruction = "Now let's learn how to open applications. Try saying 'Open Notepad' to launch the text editor.",
                    ExpectedCommands = new[] { "open notepad", "launch notepad", "start notepad", "открой блокнот" },
                    SuccessMessage = "Excellent! Notepad is now open. You can open any application this way - just say 'Open' followed by the app name.",
                    Hint = "Say 'Open Notepad' to open the text editor."
                },
                new TutorialStep
                {
                    Title = "Typing Text",
                    Instruction = "With Notepad open, let's type some text. Say 'Type: Hello World' to write text.",
                    ExpectedCommands = new[] { "type" },
                    SuccessMessage = "Perfect! You've just dictated text using your voice. This works in any text field!",
                    Hint = "Say 'Type:' followed by what you want to write."
                },
                new TutorialStep
                {
                    Title = "Selecting Text",
                    Instruction = "Now let's select the text you typed. Say 'Select all' to highlight everything.",
                    ExpectedCommands = new[] { "select all", "выдели всё" },
                    SuccessMessage = "All text is now selected. You can copy, cut, or format selected text.",
                    Hint = "Say 'Select all' to highlight all text."
                },
                new TutorialStep
                {
                    Title = "Formatting Text",
                    Instruction = "With text selected, let's make it bold. Say 'Bold' to apply bold formatting. Note: This works best in Word.",
                    ExpectedCommands = new[] { "bold", "make bold", "жирный" },
                    SuccessMessage = "The text is now bold! You can also say 'Italic' or 'Underline' for other formatting.",
                    Hint = "Say 'Bold' to make the selected text bold."
                },
                new TutorialStep
                {
                    Title = "Copying and Pasting",
                    Instruction = "Let's learn clipboard operations. First select text, then say 'Copy' to copy it.",
                    ExpectedCommands = new[] { "copy", "скопируй" },
                    SuccessMessage = "Text copied to clipboard! Now you can paste it anywhere by saying 'Paste'.",
                    Hint = "Say 'Copy' to copy selected text."
                },
                new TutorialStep
                {
                    Title = "Undoing Actions",
                    Instruction = "Made a mistake? No problem! Say 'Undo' to reverse your last action.",
                    ExpectedCommands = new[] { "undo", "отмени" },
                    SuccessMessage = "Action undone! You can say 'Undo' multiple times to go back further.",
                    Hint = "Say 'Undo' to reverse the last action."
                },
                new TutorialStep
                {
                    Title = "Saving Documents",
                    Instruction = "Always save your work! Say 'Save' to save the current document.",
                    ExpectedCommands = new[] { "save", "сохрани" },
                    SuccessMessage = "Document saved! Always remember to save your important work.",
                    Hint = "Say 'Save' to save the document."
                },
                new TutorialStep
                {
                    Title = "Getting Help",
                    Instruction = "You can ask 'How do I copy text?' or similar questions for specific instructions.",
                    ExpectedCommands = new[] { "help", "помощь", "how do", "как" },
                    SuccessMessage = "Great! You can ask questions anytime you need help with a task.",
                    Hint = "Ask 'How do I...?' for instructions."
                },
                new TutorialStep
                {
                    Title = "Tutorial Complete!",
                    Instruction = "Congratulations! You've completed the tutorial. You now know the basics of voice control. " +
                                  "Remember: Open apps, Type text, Select and format, Copy and paste, Undo mistakes, Save your work. " +
                                  "Happy computing!",
                    ExpectedCommands = Array.Empty<string>(),
                    SuccessMessage = "Tutorial complete! You're ready to use AI Companion on your own.",
                    Hint = "",
                    IsFinalStep = true
                }
            };
        }

        public void StartTutorial()
        {
            _logger?.LogInformation("Starting tutorial");
            _state = TutorialState.Active;
            _currentStepIndex = 0;

            TutorialEvent?.Invoke(this, new TutorialEventArgs
            {
                EventType = TutorialEventType.Started,
                Message = "Tutorial started"
            });

            // Announce the first step
            AnnounceCurrentStep();
        }

        public void StopTutorial()
        {
            _logger?.LogInformation("Stopping tutorial");
            _state = TutorialState.Inactive;
            _currentStepIndex = 0;

            TutorialEvent?.Invoke(this, new TutorialEventArgs
            {
                EventType = TutorialEventType.Stopped,
                Message = "Tutorial stopped"
            });

            SpeechRequested?.Invoke(this, "Tutorial stopped. Say 'Start tutorial' anytime to begin again.");
        }

        public bool ProcessCommand(string command)
        {
            // Check if the command matches what we're expecting in the current step
            if (_state != TutorialState.Active)
                return false;

            var currentStep = _tutorialSteps[_currentStepIndex];
            
            if (currentStep.IsFinalStep)
            {
                // Tutorial complete
                CompleteTutorial();
                return true;
            }

            var commandLower = command.ToLowerInvariant();
            foreach (var expected in currentStep.ExpectedCommands)
            {
                if (commandLower.Contains(expected.ToLowerInvariant()))
                {
                    // Command matched! Advance to next step
                    _logger?.LogInformation("Tutorial step {Step} completed", _currentStepIndex + 1);
                    
                    // Celebrate success
                    SpeechRequested?.Invoke(this, currentStep.SuccessMessage);

                    TutorialEvent?.Invoke(this, new TutorialEventArgs
                    {
                        EventType = TutorialEventType.StepCompleted,
                        StepIndex = _currentStepIndex,
                        Message = currentStep.SuccessMessage
                    });

                    // Move to next step after a brief pause
                    _currentStepIndex++;
                    
                    if (_currentStepIndex < _tutorialSteps.Count)
                    {
                        // Small delay before next instruction
                        System.Threading.Tasks.Task.Delay(2000).ContinueWith(_ =>
                        {
                            AnnounceCurrentStep();
                        });
                    }
                    
                    return true;
                }
            }

            return false;
        }

        public void RequestHint()
        {
            if (_state != TutorialState.Active)
                return;

            var currentStep = _tutorialSteps[_currentStepIndex];
            if (!string.IsNullOrEmpty(currentStep.Hint))
            {
                SpeechRequested?.Invoke(this, $"Hint: {currentStep.Hint}");
            }
        }

        public void SkipStep()
        {
            if (_state != TutorialState.Active)
                return;

            _logger?.LogInformation("Skipping tutorial step {Step}", _currentStepIndex + 1);
            _currentStepIndex++;

            if (_currentStepIndex >= _tutorialSteps.Count)
            {
                CompleteTutorial();
            }
            else
            {
                AnnounceCurrentStep();
            }
        }

        private void AnnounceCurrentStep()
        {
            if (_currentStepIndex >= _tutorialSteps.Count)
                return;

            var step = _tutorialSteps[_currentStepIndex];
            
            TutorialEvent?.Invoke(this, new TutorialEventArgs
            {
                EventType = TutorialEventType.StepStarted,
                StepIndex = _currentStepIndex,
                StepTitle = step.Title,
                Message = step.Instruction
            });

            SpeechRequested?.Invoke(this, step.Instruction);
        }

        private void CompleteTutorial()
        {
            _logger?.LogInformation("Tutorial completed");
            _state = TutorialState.Completed;

            TutorialEvent?.Invoke(this, new TutorialEventArgs
            {
                EventType = TutorialEventType.Completed,
                Message = "Congratulations! Tutorial complete!"
            });

            SpeechRequested?.Invoke(this, 
                "Congratulations! You've completed the tutorial and learned the basics of voice control. " +
                "You're now ready to use AI Companion on your own. Say 'Help' anytime if you need assistance.");
        }

        public TutorialStep? GetCurrentStep()
        {
            if (_currentStepIndex < _tutorialSteps.Count)
                return _tutorialSteps[_currentStepIndex];
            return null;
        }

        public (int Current, int Total) GetProgress()
        {
            return (_currentStepIndex + 1, _tutorialSteps.Count);
        }
    }

    public class TutorialStep
    {
        public string Title { get; set; } = "";
        public string Instruction { get; set; } = "";
        public string[] ExpectedCommands { get; set; } = Array.Empty<string>();
        public string SuccessMessage { get; set; } = "";
        public string Hint { get; set; } = "";
        public bool IsFinalStep { get; set; }
    }

    public enum TutorialState
    {
        Inactive,
        Active,
        Completed
    }

    public enum TutorialEventType
    {
        Started,
        StepStarted,
        StepCompleted,
        Completed,
        Stopped
    }

    public class TutorialEventArgs : EventArgs
    {
        public TutorialEventType EventType { get; set; }
        public int StepIndex { get; set; }
        public string? StepTitle { get; set; }
        public string Message { get; set; } = "";
    }
}
