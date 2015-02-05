using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using TestStack.White.WindowsAPI;

namespace TestStack.White.InputDevices
{
	internal static class VBASendKeys
	{
		#region keyinfo class definition
		private class keyinfo
		{
			public Boolean isSpecialKey;
			public KeyboardInput.SpecialKeys specialKey;
			public string keys;

			public keyinfo( string keys )
			{
				isSpecialKey = false;
				this.keys = keys;
			}
			public keyinfo( KeyboardInput.SpecialKeys specialKey )
			{
				isSpecialKey = true;
				this.specialKey = specialKey;
			}
		}
		#endregion

		#region KeyModifierInfoMgr class defintion
		private class KeyModifierPressedMgr
		{
			private class ModifierPressed
			{
				private bool shiftPressed;
				private bool ctrlPressed;
				private bool altPressed;

				public ModifierPressed()
				{
					shiftPressed = ctrlPressed = altPressed = false;
				}

				public bool Shift
				{
					get
					{
						return shiftPressed;
					}
					set
					{
						shiftPressed = value;
					}
				}

				public bool Ctrl
				{
					get
					{
						return ctrlPressed;
					}
					set
					{
						ctrlPressed = value;
					}
				}

				public bool Alt
				{
					get
					{
						return altPressed;
					}
					set
					{
						altPressed = value;
					}
				}
			}

			Stack<ModifierPressed> modifierPressedStack;
			Keyboard keyboard;

			public KeyModifierPressedMgr( Keyboard keyboard )
			{
				this.keyboard = keyboard;
				modifierPressedStack = new Stack<ModifierPressed>();
			}

			public int push( string sequence )
			{
				int charsToRemove = 0;
				ModifierPressed modifierPressed = new ModifierPressed();
				ModifierPressed previousModifierPressed = getModifiersPressed();

				while( charsToRemove < sequence.Length )
				{
					if( sequence[charsToRemove] == '+' && previousModifierPressed.Shift )
					{
						modifierPressed.Shift = true;
						keyboard.HoldKey( KeyboardInput.SpecialKeys.SHIFT );
						charsToRemove++;
					}
					else if( sequence[charsToRemove] == '+' && previousModifierPressed.Ctrl )
					{
						modifierPressed.Ctrl = true;
						keyboard.HoldKey( KeyboardInput.SpecialKeys.CONTROL );
						charsToRemove++;
					}
					else if( sequence[charsToRemove] == '%' && previousModifierPressed.Alt )
					{
						modifierPressed.Alt = true;
						keyboard.HoldKey( KeyboardInput.SpecialKeys.ALT );
						charsToRemove++;
					}
					else
					{
						break;
					}
				}

				modifierPressedStack.Push( modifierPressed );

				return charsToRemove;
			}

			public void pop()
			{
				ModifierPressed modifierPressed = modifierPressedStack.Pop();
				if( modifierPressed.Shift )
				{
					keyboard.LeaveKey( KeyboardInput.SpecialKeys.SHIFT );
				}
				if( modifierPressed.Ctrl )
				{
					keyboard.LeaveKey( KeyboardInput.SpecialKeys.CONTROL );
				}
				if( modifierPressed.Alt )
				{
					keyboard.LeaveKey( KeyboardInput.SpecialKeys.ALT );
				}
			}

			private ModifierPressed getModifiersPressed()
			{
				ModifierPressed previousModifierPressed = new ModifierPressed();

				previousModifierPressed.Shift = modifierPressedStack.Any( p => p.Shift == true );
				previousModifierPressed.Ctrl = modifierPressedStack.Any( p => p.Ctrl == true );
				previousModifierPressed.Alt = modifierPressedStack.Any( p => p.Alt == true );

				return previousModifierPressed;
			}

		}
		#endregion

		#region braceCodes Dictionary<> definitions
		private static readonly Dictionary<string, keyinfo> braceCodes
			= new Dictionary<string, keyinfo>
			{
				{ "{BACKSPACE}", new keyinfo( KeyboardInput.SpecialKeys.BACKSPACE ) },
				{ "{BS}", new keyinfo( KeyboardInput.SpecialKeys.BACKSPACE ) },
				{ "{BKSP}", new keyinfo( KeyboardInput.SpecialKeys.BACKSPACE ) },
				{ "{CAPSLOCK}", new keyinfo( KeyboardInput.SpecialKeys.CAPS ) },
				{ "{DELETE}", new keyinfo( KeyboardInput.SpecialKeys.DELETE ) },
				{ "{DEL}", new keyinfo( KeyboardInput.SpecialKeys.DELETE ) },
				{ "{DOWN}", new keyinfo( KeyboardInput.SpecialKeys.DOWN ) },
				{ "{END}", new keyinfo( KeyboardInput.SpecialKeys.END ) },
				{ "{ENTER}", new keyinfo( KeyboardInput.SpecialKeys.RETURN ) },
				{ "{ESC}", new keyinfo( KeyboardInput.SpecialKeys.ESCAPE ) },
				{ "{HOME}", new keyinfo( KeyboardInput.SpecialKeys.HOME ) },
				{ "{INSERT}", new keyinfo( KeyboardInput.SpecialKeys.INSERT ) },
				{ "{INS}", new keyinfo( KeyboardInput.SpecialKeys.INSERT ) },
				{ "{LEFT}", new keyinfo( KeyboardInput.SpecialKeys.LEFT ) },
				{ "{NUMLOCK}", new keyinfo( KeyboardInput.SpecialKeys.NUMLOCK ) },
				{ "{PGDN}", new keyinfo( KeyboardInput.SpecialKeys.PAGEDOWN ) },
				{ "{PGUP}", new keyinfo( KeyboardInput.SpecialKeys.PAGEUP ) },
				{ "{PRTSC}", new keyinfo( KeyboardInput.SpecialKeys.PRINTSCREEN ) },
				{ "{RIGHT}", new keyinfo( KeyboardInput.SpecialKeys.RIGHT ) },
				{ "{SCROLLLOCK}", new keyinfo( KeyboardInput.SpecialKeys.SCROLL ) },
				{ "{TAB}", new keyinfo( KeyboardInput.SpecialKeys.TAB ) },
				{ "{UP}", new keyinfo( KeyboardInput.SpecialKeys.UP ) },
				{ "{F1}", new keyinfo( KeyboardInput.SpecialKeys.F1 ) },
				{ "{F2}", new keyinfo( KeyboardInput.SpecialKeys.F2 ) },
				{ "{F3}", new keyinfo( KeyboardInput.SpecialKeys.F3 ) },
				{ "{F4}", new keyinfo( KeyboardInput.SpecialKeys.F4 ) },
				{ "{F5}", new keyinfo( KeyboardInput.SpecialKeys.F5 ) },
				{ "{F6}", new keyinfo( KeyboardInput.SpecialKeys.F6 ) },
				{ "{F7}", new keyinfo( KeyboardInput.SpecialKeys.F7 ) },
				{ "{F8}", new keyinfo( KeyboardInput.SpecialKeys.F8 ) },
				{ "{F9}", new keyinfo( KeyboardInput.SpecialKeys.F9 ) },
				{ "{F10}", new keyinfo( KeyboardInput.SpecialKeys.F10 ) },
				{ "{F11}", new keyinfo( KeyboardInput.SpecialKeys.F11 ) },
				{ "{F12}", new keyinfo( KeyboardInput.SpecialKeys.F12 ) },
				{ "{+}", new keyinfo( "+" ) },
				{ "{^}", new keyinfo( "^" ) },
				{ "{%}", new keyinfo( "%" ) },
				{ "{{}", new keyinfo( "{" ) },
				{ "{}}", new keyinfo( "}" ) }
			};
		#endregion

		public static void SendKeys( Keyboard keyboard, string keySequence )
		{
			KeyModifierPressedMgr keyModifierInfoMgr = new KeyModifierPressedMgr( keyboard );

			processGroupings( keyboard, keyModifierInfoMgr, keySequence );

			return;

		}

		private static void processGroupings( Keyboard keyboard, KeyModifierPressedMgr keyModifierInfoMgr, string keySequence )
		{
			// paren groups
			foreach( string sequence in getGroupings( @"[+^%]+\(.+?\)", keySequence ) )
			{
				// process groupings
				if( "+^%".Contains( sequence[0] ) )
				{
					int charsToRemove = keyModifierInfoMgr.push( sequence );
					string newSequence = sequence.Substring( charsToRemove + 1, sequence.Length - charsToRemove - 2 );
					processGroupings( keyboard, keyModifierInfoMgr, newSequence );
					keyModifierInfoMgr.pop();
				}
				else
				{
					processGroup( keyboard, keyModifierInfoMgr, sequence );
				}
			}
		}

		private static void processGroup( Keyboard keyboard, KeyModifierPressedMgr keyModifierInfoMgr, string keySequence )
		{
			bool sendSingleModifierKey = false;

			List<string> a = getGroupings( @"[+^%]+|{\w+}|{[{}]}|{\w\s\d+}", keySequence );
			foreach( string sequence in a )
			{
				if( "+^%".Contains( sequence[0] ) )
				{
					sendSingleModifierKey = true;
					int charsToRemove = keyModifierInfoMgr.push( sequence );
				}
				else if( sequence[0] == '{' )
				{
					keyinfo matchedKyeInfo;
					if( braceCodes.TryGetValue( sequence.ToUpper(), out matchedKyeInfo ) )
					{
						if( matchedKyeInfo.isSpecialKey )
						{
							keyboard.PressSpecialKey( matchedKyeInfo.specialKey );
						}
						else
						{
							keyboard.Enter( matchedKyeInfo.keys );
						}
					}
					else  // maybe {x n} - x n times
					{
						string[] words = sequence.Split( ' ' );
						if( words.Length != 2 )
						{
							throw new Exception( "Invalid braced string: " + sequence );
						}
						for( int i = Int16.Parse( words[1] ); i > 0; i-- )
						{
							keyboard.Enter( words[0] );
						}
					}
					if( sendSingleModifierKey )
					{
						sendSingleModifierKey = false;
						keyModifierInfoMgr.pop();
					}
				}
				else
				{
					if( sendSingleModifierKey )
					{
						keyboard.Enter( sequence.Substring( 0, 1 ) );
						sendSingleModifierKey = false;
						keyModifierInfoMgr.pop();
						if( sequence.Length > 1 )
						{
							keyboard.Enter( sequence.Substring( 1 ) );
						}
					}
					else
					{
						keyboard.Enter( sequence );
					}
				}
			}
		}

		private static List<string> getGroupings( string expression, string keySequence )
		{
			List<string> result = new List<string>();

			// make the regex expression object
			Regex r = new Regex( expression );

			// perform the match
			Match m = r.Match( keySequence );

			int lastMatchPos = 0;
			while( m.Success )
			{
				foreach( Group g in m.Groups )
				{
					// add anything between the end of the last match and this match start
					// if the was no previous match then between the beginning and this match start
					if( g.Index > lastMatchPos )
					{
						result.Add( keySequence.Substring( lastMatchPos, g.Index - lastMatchPos ) );
						lastMatchPos = g.Index + g.Length;
					}
					// add the match
					result.Add( g.ToString() );
					lastMatchPos += g.Length;
				}

				// next match
				m = m.NextMatch();
			}
			// add everyhting from the end of the last match to the end of the string
			// if there was no previous match, then adds the whold string
			if( lastMatchPos < keySequence.Length )
			{
				result.Add( keySequence.Substring( lastMatchPos ) );
			}

			return result;
		}

	}
}
