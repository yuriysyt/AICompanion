using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using AICompanion.Desktop.Models;

namespace AICompanion.Desktop.Controls
{
    /*
        Code-behind for the AvatarControl, handling animation state transitions.
        
        This control manages the visual representation of the AI companion,
        switching between different animations based on the current emotional
        state. The animations provide visual feedback about what the assistant
        is doing: listening, thinking, speaking, or idle.
    */
    public partial class AvatarControl : System.Windows.Controls.UserControl
    {
        /*
            Dependency property for binding the current emotion from the ViewModel.
        */
        public static readonly DependencyProperty CurrentEmotionProperty =
            DependencyProperty.Register(
                nameof(CurrentEmotion),
                typeof(AvatarEmotion),
                typeof(AvatarControl),
                new PropertyMetadata(AvatarEmotion.Neutral, OnEmotionChanged));

        public AvatarEmotion CurrentEmotion
        {
            get => (AvatarEmotion)GetValue(CurrentEmotionProperty);
            set => SetValue(CurrentEmotionProperty, value);
        }

        private Storyboard? _currentAnimation;

        public AvatarControl()
        {
            InitializeComponent();
            Loaded += OnControlLoaded;
        }

        /*
            Starts the breathing animation when the control loads.
        */
        private void OnControlLoaded(object sender, RoutedEventArgs e)
        {
            StartAnimation("BreathingAnimation");
        }

        /*
            Handles emotion changes by transitioning to the appropriate animation.
        */
        private static void OnEmotionChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is AvatarControl control)
            {
                control.TransitionToEmotion((AvatarEmotion)e.NewValue);
            }
        }

        /*
            Transitions the avatar to display the specified emotion.
        */
        private void TransitionToEmotion(AvatarEmotion emotion)
        {
            StopCurrentAnimation();

            switch (emotion)
            {
                case AvatarEmotion.Listening:
                    ShowListeningState();
                    break;

                case AvatarEmotion.Thinking:
                    ShowThinkingState();
                    break;

                case AvatarEmotion.Speaking:
                    ShowSpeakingState();
                    break;

                case AvatarEmotion.Happy:
                    ShowHappyState();
                    break;

                case AvatarEmotion.Confused:
                    ShowConfusedState();
                    break;

                default:
                    ShowNeutralState();
                    break;
            }
        }

        private void ShowNeutralState()
        {
            ListeningIndicator.Opacity = 0;
            ThinkingSpinner.Opacity = 0;
            UpdateMouthPath("M 60,125 Q 90,145 120,125");
            StartAnimation("BreathingAnimation");
        }

        private void ShowListeningState()
        {
            ListeningIndicator.Opacity = 1;
            ThinkingSpinner.Opacity = 0;
            UpdateMouthPath("M 60,130 L 120,130");
            StartAnimation("ListeningAnimation");
        }

        private void ShowThinkingState()
        {
            ListeningIndicator.Opacity = 0;
            ThinkingSpinner.Opacity = 1;
            UpdateMouthPath("M 60,130 Q 90,125 120,130");
            StartAnimation("ThinkingAnimation");
        }

        private void ShowSpeakingState()
        {
            ListeningIndicator.Opacity = 0;
            ThinkingSpinner.Opacity = 0;
            UpdateMouthPath("M 70,120 Q 90,140 110,120");
        }

        private void ShowHappyState()
        {
            ListeningIndicator.Opacity = 0;
            ThinkingSpinner.Opacity = 0;
            UpdateMouthPath("M 55,120 Q 90,155 125,120");
            StartAnimation("BreathingAnimation");
        }

        private void ShowConfusedState()
        {
            ListeningIndicator.Opacity = 0;
            ThinkingSpinner.Opacity = 0;
            UpdateMouthPath("M 70,130 Q 90,125 110,135");
        }

        /*
            Updates the mouth shape using path data.
        */
        private void UpdateMouthPath(string pathData)
        {
            Mouth.Data = System.Windows.Media.Geometry.Parse(pathData);
        }

        /*
            Starts the specified storyboard animation.
        */
        private void StartAnimation(string animationName)
        {
            if (Resources[animationName] is Storyboard storyboard)
            {
                _currentAnimation = storyboard;
                storyboard.Begin();
            }
        }

        /*
            Stops any currently running animation.
        */
        private void StopCurrentAnimation()
        {
            _currentAnimation?.Stop();
            _currentAnimation = null;
        }
    }
}
