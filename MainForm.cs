using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SpeechLib;

namespace SAPItest.UILayer
{
	public class MainForm : Form
	{
		private TextBox _txtText;
		private Label _label1;
		private Label _label2;
		private Label _label3;
		private Button _btnSpeak;
		private ListBox _voicesListBox;
		private SpVoice _speech;
		private ListBox _logListBox;
		private Button _btnStop;
		private Label label1;
		private Label label2;
		private Label label3;
		private Label label4;
		private TrackBar _tBarRateSapi;
		private TrackBar _tBarVolumeSapi;
		private TrackBar _tBarRateXml;
		private TrackBar _tBarVolumeXml;
		private TrackBar _tBarPitchXml;
		private Label label5;
		private Label label6;
		private readonly List<SpObjectToken> _tokens = new List<SpObjectToken>();

		private int _xmlRate = 0;
		private int _xmlVolume = 100;
		private Button _btnSpellSpeak;
		private int _xmlPitch = 0;

		public MainForm()
		{
			InitializeComponent();
		}

		private void InitializeComponent()
		{
			this._voicesListBox = new System.Windows.Forms.ListBox();
			this._txtText = new System.Windows.Forms.TextBox();
			this._label1 = new System.Windows.Forms.Label();
			this._label2 = new System.Windows.Forms.Label();
			this._label3 = new System.Windows.Forms.Label();
			this._btnSpeak = new System.Windows.Forms.Button();
			this._logListBox = new System.Windows.Forms.ListBox();
			this._btnStop = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this._tBarRateSapi = new System.Windows.Forms.TrackBar();
			this._tBarVolumeSapi = new System.Windows.Forms.TrackBar();
			this._tBarRateXml = new System.Windows.Forms.TrackBar();
			this._tBarVolumeXml = new System.Windows.Forms.TrackBar();
			this._tBarPitchXml = new System.Windows.Forms.TrackBar();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this._btnSpellSpeak = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this._tBarRateSapi)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarVolumeSapi)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarRateXml)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarVolumeXml)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarPitchXml)).BeginInit();
			this.SuspendLayout();
			// 
			// _voicesListBox
			// 
			this._voicesListBox.FormattingEnabled = true;
			this._voicesListBox.Location = new System.Drawing.Point(270, 1);
			this._voicesListBox.Name = "_voicesListBox";
			this._voicesListBox.Size = new System.Drawing.Size(616, 69);
			this._voicesListBox.TabIndex = 0;
			this._voicesListBox.SelectedIndexChanged += new System.EventHandler(this.VoicesListBoxSelectedIndexChanged);
			// 
			// _txtText
			// 
			this._txtText.Location = new System.Drawing.Point(270, 73);
			this._txtText.Multiline = true;
			this._txtText.Name = "_txtText";
			this._txtText.Size = new System.Drawing.Size(378, 253);
			this._txtText.TabIndex = 1;
			this._txtText.Text = "Text to read.";
			// 
			// _label1
			// 
			this._label1.AutoSize = true;
			this._label1.Location = new System.Drawing.Point(666, 90);
			this._label1.Name = "_label1";
			this._label1.Size = new System.Drawing.Size(33, 13);
			this._label1.TabIndex = 4;
			this._label1.Text = "Rate:";
			// 
			// _label2
			// 
			this._label2.AutoSize = true;
			this._label2.Location = new System.Drawing.Point(654, 120);
			this._label2.Name = "_label2";
			this._label2.Size = new System.Drawing.Size(45, 13);
			this._label2.TabIndex = 4;
			this._label2.Text = "Volume:";
			// 
			// _label3
			// 
			this._label3.AutoSize = true;
			this._label3.Location = new System.Drawing.Point(231, 4);
			this._label3.Name = "_label3";
			this._label3.Size = new System.Drawing.Size(37, 13);
			this._label3.TabIndex = 4;
			this._label3.Text = "Voice:";
			// 
			// _btnSpeak
			// 
			this._btnSpeak.Location = new System.Drawing.Point(715, 313);
			this._btnSpeak.Name = "_btnSpeak";
			this._btnSpeak.Size = new System.Drawing.Size(59, 23);
			this._btnSpeak.TabIndex = 5;
			this._btnSpeak.Text = "Speak";
			this._btnSpeak.UseVisualStyleBackColor = true;
			this._btnSpeak.Click += new System.EventHandler(this.BtnSpeakClick);
			// 
			// _logListBox
			// 
			this._logListBox.FormattingEnabled = true;
			this._logListBox.HorizontalScrollbar = true;
			this._logListBox.Location = new System.Drawing.Point(2, 28);
			this._logListBox.Name = "_logListBox";
			this._logListBox.Size = new System.Drawing.Size(263, 303);
			this._logListBox.TabIndex = 6;
			// 
			// _btnStop
			// 
			this._btnStop.Location = new System.Drawing.Point(657, 313);
			this._btnStop.Name = "_btnStop";
			this._btnStop.Size = new System.Drawing.Size(52, 23);
			this._btnStop.TabIndex = 7;
			this._btnStop.Text = "Stop";
			this._btnStop.UseVisualStyleBackColor = true;
			this._btnStop.Click += new System.EventHandler(this.BtnStopClick);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(658, 171);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(32, 13);
			this.label1.TabIndex = 4;
			this.label1.Text = "XML:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(656, 73);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "SAPI:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(668, 191);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 13);
			this.label3.TabIndex = 4;
			this.label3.Text = "Rate:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(656, 228);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(45, 13);
			this.label4.TabIndex = 4;
			this.label4.Text = "Volume:";
			// 
			// _tBarRateSapi
			// 
			this._tBarRateSapi.Location = new System.Drawing.Point(696, 82);
			this._tBarRateSapi.Minimum = -10;
			this._tBarRateSapi.Name = "_tBarRateSapi";
			this._tBarRateSapi.Size = new System.Drawing.Size(191, 45);
			this._tBarRateSapi.TabIndex = 8;
			this._tBarRateSapi.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// _tBarVolumeSapi
			// 
			this._tBarVolumeSapi.Location = new System.Drawing.Point(696, 122);
			this._tBarVolumeSapi.Maximum = 100;
			this._tBarVolumeSapi.Name = "_tBarVolumeSapi";
			this._tBarVolumeSapi.Size = new System.Drawing.Size(191, 45);
			this._tBarVolumeSapi.TabIndex = 8;
			this._tBarVolumeSapi.Value = 100;
			this._tBarVolumeSapi.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// _tBarRateXml
			// 
			this._tBarRateXml.Location = new System.Drawing.Point(696, 186);
			this._tBarRateXml.Minimum = -10;
			this._tBarRateXml.Name = "_tBarRateXml";
			this._tBarRateXml.Size = new System.Drawing.Size(191, 45);
			this._tBarRateXml.TabIndex = 8;
			this._tBarRateXml.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// _tBarVolumeXml
			// 
			this._tBarVolumeXml.Location = new System.Drawing.Point(696, 227);
			this._tBarVolumeXml.Maximum = 100;
			this._tBarVolumeXml.Name = "_tBarVolumeXml";
			this._tBarVolumeXml.Size = new System.Drawing.Size(191, 45);
			this._tBarVolumeXml.TabIndex = 8;
			this._tBarVolumeXml.Value = 100;
			this._tBarVolumeXml.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// _tBarPitchXml
			// 
			this._tBarPitchXml.Location = new System.Drawing.Point(696, 265);
			this._tBarPitchXml.Minimum = -10;
			this._tBarPitchXml.Name = "_tBarPitchXml";
			this._tBarPitchXml.Size = new System.Drawing.Size(191, 45);
			this._tBarPitchXml.TabIndex = 8;
			this._tBarPitchXml.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(667, 264);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 13);
			this.label5.TabIndex = 4;
			this.label5.Text = "Pitch:";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(1, 12);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(28, 13);
			this.label6.TabIndex = 4;
			this.label6.Text = "Log:";
			// 
			// _btnSpellSpea
			// 
			this._btnSpellSpeak.Location = new System.Drawing.Point(780, 313);
			this._btnSpellSpeak.Name = "_btnSpellSpea";
			this._btnSpellSpeak.Size = new System.Drawing.Size(101, 23);
			this._btnSpellSpeak.TabIndex = 5;
			this._btnSpellSpeak.Text = "Spell speak";
			this._btnSpellSpeak.UseVisualStyleBackColor = true;
			this._btnSpellSpeak.Click += new System.EventHandler(this.BtnSpellSpeakClick);
			// 
			// MainForm
			// 
			this.ClientSize = new System.Drawing.Size(893, 338);
			this.Controls.Add(this._tBarPitchXml);
			this.Controls.Add(this._tBarVolumeXml);
			this.Controls.Add(this._tBarRateXml);
			this.Controls.Add(this._tBarVolumeSapi);
			this.Controls.Add(this._tBarRateSapi);
			this.Controls.Add(this._btnStop);
			this.Controls.Add(this._logListBox);
			this.Controls.Add(this._btnSpellSpeak);
			this.Controls.Add(this._btnSpeak);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this._label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this._label3);
			this.Controls.Add(this.label3);
			this.Controls.Add(this._label1);
			this.Controls.Add(this._txtText);
			this.Controls.Add(this._voicesListBox);
			this.Name = "MainForm";
			this.Text = "SAPI tester";
			this.Load += new System.EventHandler(this.MainFormLoad);
			((System.ComponentModel.ISupportInitialize)(this._tBarRateSapi)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarVolumeSapi)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarRateXml)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarVolumeXml)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this._tBarPitchXml)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		private void MainFormLoad(object sender, EventArgs e)
		{
			_speech = new SpVoice();
			var tokens = _speech.GetVoices();
			for (var i = 0; i < tokens.Count; i++)
			{
				_tokens.Add(tokens.Item(i));
				_voicesListBox.Items.Add(tokens.Item(i).GetDescription());
				if (i != 0) continue;
				_speech.Voice = tokens.Item(0);
				_voicesListBox.SelectedIndex = 0;
			}
		}

		private void BtnSpeakClick(object sender, EventArgs e)
		{
			_logListBox.Items.Add("speak");
			try
			{
				_speech.Speak(
					string.Format("<rate absspeed=\"{0}\"><volume level=\"{1}\"><pitch absmiddle=\"{2}\">{3}</pitch></volume></rate>",
								  _xmlRate, _xmlVolume, _xmlPitch, _txtText.Text),
					SpeechVoiceSpeakFlags.SVSFlagsAsync | SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);
			}
			catch (Exception)
			{
				_logListBox.Items.Add("speak error");
			}
			ClearLog();
		}

		private void VoicesListBoxSelectedIndexChanged(object sender, EventArgs e)
		{
			_speech.Voice = _tokens[_voicesListBox.SelectedIndex];
		}

		private void BtnStopClick(object sender, EventArgs e)
		{
			_logListBox.Items.Add("stop");
			try
			{
				_speech.Speak("", SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);
			}
			catch
			{
			}
			ClearLog();
		}

		private void ClearLog()
		{
			if (_logListBox.Items.Count > 13)
				for (var i = 0; i < _logListBox.Items.Count - 13; i++) _logListBox.Items.RemoveAt(0);
		}

		private void TrackBarScroll(object sender, EventArgs e)
		{
			var item = (TrackBar)sender;
			switch (item.Name)
			{
				case "_tBarRateSapi":
					_logListBox.Items.Add("set sapi voice rate = " + item.Value);
					try
					{
						_speech.Rate = Convert.ToInt32(item.Value);
					}
					catch (Exception ex)
					{
						_logListBox.Items.Add(ex.Message);
					}
					break;
				case "_tBarVolumeSapi":
					_logListBox.Items.Add("set sapi voice volume = " + item.Value);
					try
					{
						_speech.Volume = Convert.ToInt32(item.Value);
					}
					catch (Exception ex)
					{
						_logListBox.Items.Add(ex.Message);
					}
					break;
				case "_tBarRateXml":
					_logListBox.Items.Add("set xml rate = " + item.Value);
					try
					{
						_xmlRate = Convert.ToInt32(item.Value);
					}
					catch (Exception ex)
					{
						_logListBox.Items.Add(ex.Message);
					}
					break;
				case "_tBarVolumeXml":
					_logListBox.Items.Add("set xml voice volume=" + item.Value);
					try
					{
						_xmlVolume = Convert.ToInt32(item.Value);
					}
					catch (Exception ex)
					{
						_logListBox.Items.Add(ex.Message);
					}
					break;
				case "_tBarPitchXml":
					_logListBox.Items.Add("set xml pitch = " + item.Value);
					try
					{
						_xmlPitch = Convert.ToInt32(item.Value);
					}
					catch (Exception ex)
					{
						_logListBox.Items.Add(ex.Message);
					}
					break;
			}
			ClearLog();
		}

		private void BtnSpellSpeakClick(object sender, EventArgs e)
		{
			_logListBox.Items.Add("speak");
			try
			{
				_speech.Speak(
					string.Format(
						"<spell><rate absspeed=\"{0}\"><volume level=\"{1}\"><pitch absmiddle=\"{2}\">{3}</pitch></volume></rate></spell>",
						_xmlRate, _xmlVolume, _xmlPitch, _txtText.Text),
					SpeechVoiceSpeakFlags.SVSFlagsAsync | SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);
			}
			catch (Exception)
			{
				_logListBox.Items.Add("speak error");
			}
			ClearLog();
		}
	}
}
