#include "stdafx.h"
#include "TranscribeDlg.h"
#include "Utils/logger.h"
#include <regex>
#include <iostream>
#include <filesystem>

HRESULT TranscribeDlg::show()
{
	auto res = DoModal( nullptr );
	if( res == -1 )
		return HRESULT_FROM_WIN32( GetLastError() );
	switch( res )
	{
	case IDC_BACK:
		return SCREEN_MODEL;
	case IDC_CAPTURE:
		return SCREEN_CAPTURE;
	}
	return S_OK;
}

constexpr int progressMaxInteger = 1024 * 8;

static const LPCTSTR regValInput = L"sourceMedia";
static const LPCTSTR regValOutFormat = L"resultFormat";
static const LPCTSTR regValOutPath = L"resultPath";
static const LPCTSTR regValUseInputFolder = L"useInputFolder";

LRESULT TranscribeDlg::OnInitDialog( UINT nMessage, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// First DDX call, hooks up variables to controls.
	DoDataExchange( false );
	printModelDescription();
	languageSelector.initialize( m_hWnd, IDC_LANGUAGE, appState );
	cbConsole.initialize( m_hWnd, IDC_CONSOLE, appState );
	cbTranslate.initialize( m_hWnd, IDC_TRANSLATE, appState );
	populateOutputFormats();

	pendingState.initialize(
		{
			languageSelector, GetDlgItem( IDC_TRANSLATE ),
			sourceMediaPath, GetDlgItem( IDC_BROWSE_MEDIA ),
			transcribeOutFormat, useInputFolder,
			transcribeOutputPath, GetDlgItem( IDC_BROWSE_RESULT ),
			GetDlgItem( IDCANCEL ),
			GetDlgItem( IDC_BACK ),
			GetDlgItem( IDC_CAPTURE )
		},
		{
			progressBar, GetDlgItem( IDC_PENDING_TEXT )
		} );

	HRESULT hr = work.create( this );
	if( FAILED( hr ) )
	{
		reportError( m_hWnd, L"CreateThreadpoolWork failed", nullptr, hr );
		EndDialog( IDCANCEL );
	}

	progressBar.SetRange32( 0, progressMaxInteger );
	progressBar.SetStep( 1 );

	sourceMediaPath.SetWindowText( appState.stringLoad( regValInput ) );
	transcribeOutFormat.SetCurSel( (int)appState.dwordLoad( regValOutFormat, 0 ) );
	transcribeOutputPath.SetWindowText( appState.stringLoad( regValOutPath ) );
	if( appState.boolLoad( regValUseInputFolder ) )
		useInputFolder.SetCheck( BST_CHECKED );
	BOOL unused;
	onOutFormatChange( 0, 0, nullptr, unused );

	appState.lastScreenSave( SCREEN_TRANSCRIBE );
	appState.setupIcon( this );
	ATLVERIFY( CenterWindow() );
	return 0;
}

void TranscribeDlg::printModelDescription()
{
	CString text;
	if( S_OK == appState.model->isMultilingual() )
		text = L"Multilingual";
	else
		text = L"Single-language";
	text += L" model \"";
	LPCTSTR path = appState.source.path;
	path = ::PathFindFileName( path );
	text += path;
	text += L"\", ";
	const int64_t cb = appState.source.sizeInBytes;
	if( cb < 1 << 30 )
	{
		constexpr double mul = 1.0 / ( 1 << 20 );
		double mb = (double)cb * mul;
		text.AppendFormat( L"%.1f MB", mb );
	}
	else
	{
		constexpr double mul = 1.0 / ( 1 << 30 );
		double gb = (double)cb * mul;
		text.AppendFormat( L"%.2f GB", gb );
	}
	text += L" on disk, ";
	text += implString( appState.source.impl );
	text += L" implementation";

	modelDesc.SetWindowText( text );
}

// Populate the "Output Format" combobox
void TranscribeDlg::populateOutputFormats()
{
	transcribeOutFormat.AddString( L"None" );
	transcribeOutFormat.AddString( L"Text file" );
	transcribeOutFormat.AddString( L"Text with timestamps" );
	transcribeOutFormat.AddString( L"SubRip subtitles" );
	transcribeOutFormat.AddString( L"WebVTT subtitles" );
}

// The enum values should match 0-based indices of the combobox items
enum struct TranscribeDlg::eOutputFormat : uint8_t
{
	None = 0,
	Text = 1,
	TextTimestamps = 2,
	SubRip = 3,
	WebVTT = 4,
};

enum struct TranscribeDlg::eVisualState : uint8_t
{
	Idle = 0,
	Running = 1,
	Stopping = 2
};

// CBN_SELCHANGE notification for IDC_OUTPUT_FORMAT combobox
LRESULT TranscribeDlg::onOutFormatChange( UINT, INT, HWND, BOOL& bHandled )
{
	BOOL enabled = transcribeOutFormat.GetCurSel() != 0;
	useInputFolder.EnableWindow( enabled );

	if( isChecked( useInputFolder ) && enabled )
	{
		enabled = FALSE;
		setOutputPath();
	}
	transcribeOutputPath.EnableWindow( enabled );
	transcribeOutputBrowse.EnableWindow( enabled );

	return 0;
}

// EN_CHANGE notification for IDC_PATH_MEDIA edit box
LRESULT TranscribeDlg::onInputChange( UINT, INT, HWND, BOOL& )
{
	if( !useInputFolder.IsWindowEnabled() )
		return 0;
	if( !isChecked( useInputFolder ) )
		return 0;
	setOutputPath();
	return 0;
}

void TranscribeDlg::onBrowseMedia()
{
	LPCTSTR title = L"Input audio file to transcribe";
	LPCTSTR filters = L"Multimedia Files\0*.wav;*.wave;*.mp3;*.wma;*.mp4;*.mpeg4;*.mkv;*.m4a\0\0";

	CString path;
	sourceMediaPath.GetWindowText( path );
	if( !getOpenFileName( m_hWnd, title, filters, path ) )
		return;
	sourceMediaPath.SetWindowText( path );
	if( useInputFolder.IsWindowEnabled() && useInputFolder.GetCheck() == BST_CHECKED )
		setOutputPath( path );
}

static const LPCTSTR outputFilters = L"Text files (*.txt)\0*.txt\0Text with timestamps (*.txt)\0*.txt\0SubRip subtitles (*.srt)\0*.srt\0WebVTT subtitles (*.vtt)\0*.vtt\0\0";
static const std::array<LPCTSTR, 4> outputExtensions =
{
	L".txt", L".txt", L".srt", L".vtt"
};

void TranscribeDlg::setOutputPath( const CString& input )
{
	const int format = transcribeOutFormat.GetCurSel() - 1;
	if( format < 0 || format >= outputExtensions.size() )
		return;
	const LPCTSTR ext = outputExtensions[ format ];
	CString path = input;
	path.Trim();
	const bool renamed = PathRenameExtension( path.GetBufferSetLength( path.GetLength() + 4 ), ext );
	path.ReleaseBuffer();
	if( !renamed )
		return;
	transcribeOutputPath.SetWindowText( path );
}

void TranscribeDlg::setOutputPath()
{
	CString path;
	if( !sourceMediaPath.GetWindowText( path ) )
		return;
	if( path.GetLength() <= 0 )
		return;
	setOutputPath( path );
}

void TranscribeDlg::onInputFolderCheck()
{
	const bool checked = isChecked( useInputFolder );

	BOOL enableOutput = checked ? FALSE : TRUE;
	transcribeOutputPath.EnableWindow( enableOutput );
	transcribeOutputBrowse.EnableWindow( enableOutput );

	if( !checked )
		return;
	setOutputPath();
}

void TranscribeDlg::onBrowseOutput()
{
	const DWORD origFilterIndex = (DWORD)transcribeOutFormat.GetCurSel() - 1;

	LPCTSTR title = L"Output Text File";
	CString path;
	transcribeOutputPath.GetWindowText( path );
	DWORD filterIndex = origFilterIndex;
	if( !getSaveFileName( m_hWnd, title, outputFilters, path, &filterIndex ) )
		return;

	LPCTSTR ext = PathFindExtension( path );
	if( 0 == *ext && filterIndex < outputExtensions.size() )
	{
		wchar_t* const buffer = path.GetBufferSetLength( path.GetLength() + 5 );
		PathRenameExtension( buffer, outputExtensions[ filterIndex ] );
		path.ReleaseBuffer();
	}

	transcribeOutputPath.SetWindowText( path );
	if( filterIndex != origFilterIndex )
		transcribeOutFormat.SetCurSel( filterIndex + 1 );
}

void TranscribeDlg::setPending( bool nowPending )
{
	pendingState.setPending( nowPending );
}

void TranscribeDlg::transcribeError( LPCTSTR text, HRESULT hr )
{
	reportError( m_hWnd, text, L"Unable to transcribe audio", hr );
}

std::string utf8(const std::wstring& utf16)
{
	int count = WideCharToMultiByte(CP_UTF8, 0, utf16.c_str(), (int)utf16.length(), nullptr, 0, nullptr, nullptr);
	std::string str(count, 0);
	WideCharToMultiByte(CP_UTF8, 0, utf16.c_str(), -1, &str[0], count, nullptr, nullptr);
	return str;
}

void TranscribeDlg::onTranscribe()
{
	switch( transcribeArgs.visualState )
	{
	case eVisualState::Running:
		transcribeArgs.visualState = eVisualState::Stopping;
		transcribeButton.EnableWindow( FALSE );
		return;
	case eVisualState::Stopping:
		return;
	}

	// Validate input
	sourceMediaPath.GetWindowText( transcribeArgs.pathMedia );
	if( transcribeArgs.pathMedia.GetLength() <= 0 )
	{
		transcribeError( L"Please select an input audio file" );
		return;
	}

	if( !PathFileExists( transcribeArgs.pathMedia ) )
	{
		transcribeError( L"Input audio file does not exist", HRESULT_FROM_WIN32( ERROR_FILE_NOT_FOUND ) );
		return;
	}

	transcribeArgs.language = languageSelector.selectedLanguage();
	transcribeArgs.translate = cbTranslate.checked();
	if( isInvalidTranslate( m_hWnd, transcribeArgs.language, transcribeArgs.translate ) )
		return;

	transcribeArgs.format = (eOutputFormat)(uint8_t)transcribeOutFormat.GetCurSel();
	if( transcribeArgs.format != eOutputFormat::None )
	{
		transcribeStartTime.GetWindowText(transcribeArgs.customStartTimeText);
		transcribeEndTime.GetWindowText(transcribeArgs.customEndTimeText);
		transcribeOutputPath.GetWindowText( transcribeArgs.pathOutput );
		if( transcribeArgs.pathOutput.GetLength() <= 0 )
		{
			transcribeError( L"Please select an output text file" );
			return;
		}
		if (PathFileExists(transcribeArgs.pathOutput))
		{
			const int resp = MessageBox(L"The output file is already there.\nadd subfix to file name?", L"Confirm New File", MB_ICONQUESTION | MB_YESNO);
			if (resp == IDYES) {
				std::wstring myPath((LPCTSTR)transcribeArgs.pathOutput);
				std::wstring fileName = std::filesystem::path(myPath).filename();
				std::wstring fileExtension = std::filesystem::path(myPath).extension();
				std::wregex pattern(LR"(.*\((\d+)\)\.(.*))");
			
				std::wsmatch match;
				if (std::regex_match(fileName, match, pattern)) {
					int number = std::stoi(match.str(1));

					CString oldFileExtension;
					oldFileExtension.Format(_T("(%d)%s"), number, fileExtension.c_str());

					CString newFileExtension;
					newFileExtension.Format(_T("(%d)%s"), number+1, fileExtension.c_str());

					transcribeArgs.pathOutput.Replace(oldFileExtension, newFileExtension);
					transcribeOutputPath.SetWindowText(transcribeArgs.pathOutput);
				} else {
					CString newFileExtension;
					newFileExtension.Format(_T("(2)%s"), fileExtension.c_str());
					transcribeArgs.pathOutput.Replace(fileExtension.c_str(), newFileExtension);
					transcribeOutputPath.SetWindowText(transcribeArgs.pathOutput);
				}
			} else {
				const int resp = MessageBox(L"The output file is already there.\nOverwrite the file?", L"Confirm Overwrite", MB_ICONQUESTION | MB_YESNO);
				if (resp != IDYES)
					return;
			}
		}
		appState.stringStore(regValOutPath, transcribeArgs.pathOutput);

		CString startTime = transcribeArgs.customStartTimeText;
		CString endTime = transcribeArgs.customEndTimeText;
		std::regex pattern(R"((\d+):(\d+):(\d+)\.(\d+))");
		std::smatch matches;
		std::string startTimeString = utf8(startTime.GetString());
		std::string endTimeString = utf8(endTime.GetString());
		if (std::regex_match(startTimeString, matches, pattern)) {
			int hour = std::stoi(matches[1]);
			int min = std::stoi(matches[2]);
			int sec = std::stoi(matches[3]);
			int milli = 0;
			if (matches[4].length() == 3) {
				milli = std::stoi(matches[4]);
			}
			else if (matches[4].length() == 2) {
				milli = std::stoi(matches[4]) * 10;
			}
			else if (matches[4].length() == 1) {
				milli = std::stoi(matches[4]) * 100;
			}

			transcribeArgs.customStartTimeMillis = hour * 60 * 60 * 1000 + min * 60 * 1000 + sec * 1000 + milli * 100;
		}
		else {
			transcribeArgs.customStartTimeMillis = _ttoi(startTime) * 1000;
		}

		if (std::regex_match(endTimeString, matches, pattern)) {
			int hour = std::stoi(matches[1]);
			int min = std::stoi(matches[2]);
			int sec = std::stoi(matches[3]);
			int milli = 0;
			if (matches[4].length() == 3) {
				milli = std::stoi(matches[4]);
			}
			else if (matches[4].length() == 2) {
				milli = std::stoi(matches[4]) * 10;
			}
			else if (matches[4].length() == 1) {
				milli = std::stoi(matches[4]) * 100;
			}

			transcribeArgs.customEndTimeMillis = hour * 60 * 60 * 1000 + min * 60 * 1000 + sec * 1000 + milli * 100;
		}
		else {
			transcribeArgs.customEndTimeMillis = _ttoi(endTime) * 1000;
		}
	}
	else
		cbConsole.ensureChecked();

	appState.dwordStore( regValOutFormat, (uint32_t)(int)transcribeArgs.format );
	appState.boolStore( regValUseInputFolder, isChecked( useInputFolder ) );
	languageSelector.saveSelection( appState );
	cbTranslate.saveSelection( appState );
	appState.stringStore( regValInput, transcribeArgs.pathMedia );

	setPending( true );
	transcribeArgs.visualState = eVisualState::Running;
	transcribeButton.SetWindowText( L"Stop" );
	work.post();
}

void __stdcall TranscribeDlg::poolCallback() noexcept
{
	HRESULT hr = transcribe();
	PostMessage( WM_CALLBACK_STATUS, (WPARAM)hr );
}

static void printTime( CString& rdi, int64_t ticks )
{
	const Whisper::sTimeSpan ts{ (uint64_t)ticks };
	const Whisper::sTimeSpanFields fields = ts;

	if( fields.days != 0 )
	{
		rdi.AppendFormat( L"%i days, %i hours", fields.days, (int)fields.hours );
		return;
	}
	if( ( fields.hours | fields.minutes ) != 0 )
	{
		rdi.AppendFormat( L"%02d:%02d:%02d", (int)fields.hours, (int)fields.minutes, (int)fields.seconds );
		return;
	}
	rdi.AppendFormat( L"%.3f seconds", (double)ticks / 1E7 );
}

LRESULT TranscribeDlg::onCallbackStatus( UINT, WPARAM wParam, LPARAM, BOOL& bHandled )
{
	setPending( false );
	transcribeButton.SetWindowText( L"Transcribe" );
	transcribeButton.EnableWindow( TRUE );
	const bool prematurely = ( transcribeArgs.visualState == eVisualState::Stopping );
	transcribeArgs.visualState = eVisualState::Idle;

	const HRESULT hr = (HRESULT)wParam;
	if( FAILED( hr ) )
	{
		LPCTSTR failMessage = L"Transcribe failed";

		if( transcribeArgs.errorMessage.GetLength() > 0 )
		{
			CString tmp = failMessage;
			tmp += L"\n";
			tmp += transcribeArgs.errorMessage;
			transcribeError( tmp, hr );
		}
		else
			transcribeError( failMessage, hr );

		return 0;
	}

	const int64_t elapsed = ( GetTickCount64() - transcribeArgs.startTime ) * 10'000;
	const int64_t media = transcribeArgs.mediaDuration;
	CString message;
	if( prematurely )
		message = L"Transcribed an initial portion of the audio";
	else
		message = L"Transcribed the audio";
	message += L"\nMedia duration: ";
	printTime( message, media );
	message += L"\nProcessing time: ";
	printTime( message, elapsed );
	message += L"\nRelative processing speed: ";
	double mul = (double)media / (double)elapsed;
	message.AppendFormat( L"%g", mul );

	MessageBox( message, L"Transcribe Completed", MB_OK | MB_ICONINFORMATION );
	return 0;
}

void TranscribeDlg::getThreadError()
{
	getLastError( transcribeArgs.errorMessage );
}

#define CHECK_EX( hr ) { const HRESULT __hr = ( hr ); if( FAILED( __hr ) ) { getThreadError(); return __hr; } }

HRESULT TranscribeDlg::transcribe()
{
	transcribeArgs.startTime = GetTickCount64();
	clearLastError();
	transcribeArgs.errorMessage = L"";

	using namespace Whisper;

	const eOutputFormat format = transcribeArgs.format;
	CAtlFile outputFile;
	if( format != eOutputFormat::None )
		CHECK( outputFile.Create( transcribeArgs.pathOutput, GENERIC_WRITE, 0, CREATE_ALWAYS ) );

	transcribeArgs.resultFlags = eResultFlags::Timestamps | eResultFlags::Tokens;

	CComPtr<iContext> context;
	CHECK_EX( appState.model->createContext( &context ) );

	sFullParams fullParams;
	CHECK_EX( context->fullDefaultParams( eSamplingStrategy::Greedy, &fullParams ) );
	
	fullParams.language = transcribeArgs.language;
	fullParams.setFlag( eFullParamsFlags::Translate, transcribeArgs.translate );
	//fullParams.setFlag( eFullParamsFlags::TokenTimestamps, true);
	//fullParams.setFlag( eFullParamsFlags::SingleSegment, true);
	//fullParams.setFlag( eFullParamsFlags::NoContext, true);
	fullParams.resetFlag( eFullParamsFlags::PrintRealtime );

	// Setup the callbacks
	fullParams.new_segment_callback = &newSegmentCallbackStatic;
	fullParams.new_segment_callback_user_data = this;
	fullParams.encoder_begin_callback = &encoderBeginCallback;
	fullParams.encoder_begin_callback_user_data = this;
	
	int customeStartTime = transcribeArgs.customStartTimeMillis;
	int customeEndTime = transcribeArgs.customEndTimeMillis;
	fullParams.offset_ms = customeStartTime;
	if (customeEndTime > customeStartTime) {
		fullParams.duration_ms = (customeEndTime - customeStartTime);
	}
	else {
		fullParams.duration_ms = 0;
	}

	// Setup Audio Reader
	CComPtr<iAudioReader> reader;
	CHECK_EX(appState.mediaFoundation->openAudioFile(transcribeArgs.pathMedia, false, &reader, customeStartTime));

	// Setup the progress indication sink
	sProgressSink progressSink{ &progressCallbackStatic, this };
	// Run the transcribe
	CHECK_EX( context->runStreamed( fullParams, progressSink, reader ) );

	// Once finished, query duration of the audio.
	// The duration before the processing is sometimes different, by 20 seconds for the file in that issue:
	// https://github.com/Const-me/Whisper/issues/4
	CHECK_EX( reader->getDuration( transcribeArgs.mediaDuration ) );

	// context->timingsPrint();

	if( format == eOutputFormat::None )
		return S_OK;

	CComPtr<iTranscribeResult> result;
	CHECK_EX( context->getResults( transcribeArgs.resultFlags, &result ) );

	sTranscribeLength len;
	CHECK_EX( result->getSize( len ) );
	const sSegment* const segments = result->getSegments();

	switch( format )
	{
	case eOutputFormat::Text:
		return writeTextFile( segments, len.countSegments, outputFile, false );
	case eOutputFormat::TextTimestamps:
		return writeTextFile( segments, len.countSegments, outputFile, true );
	case eOutputFormat::SubRip:
		return writeSubRip( segments, len.countSegments, appState.duplicatedResults.size(), outputFile);
	case eOutputFormat::WebVTT:
		return writeWebVTT( segments, len.countSegments, outputFile );
	default:
		return E_FAIL;
	}
}

#undef CHECK_EX

inline HRESULT TranscribeDlg::progressCallback( double p ) noexcept
{
	constexpr double mul = progressMaxInteger;
	int pos = lround( mul * p );
	progressBar.PostMessage( PBM_SETPOS, pos, 0 );
	return S_OK;
}

HRESULT __cdecl TranscribeDlg::progressCallbackStatic( double p, Whisper::iContext* ctx, void* pv ) noexcept
{
	TranscribeDlg* dlg = (TranscribeDlg*)pv;
	return dlg->progressCallback( p );
}

namespace
{
	HRESULT write( CAtlFile& file, const CStringA& line )
	{
		if( line.GetLength() > 0 )
			CHECK( file.Write( cstr( line ), (DWORD)line.GetLength() ) );
		return S_OK;
	}

	const char* skipBlank( const char* rsi )
	{
		while( true )
		{
			const char c = *rsi;
			if( c == ' ' || c == '\t' )
			{
				rsi++;
				continue;
			}
			return rsi;
		}
	}
}

using Whisper::sSegment;


HRESULT TranscribeDlg::writeTextFile( const sSegment* const segments, const size_t length, CAtlFile& file, bool timestamps )
{
	using namespace Whisper;
	CHECK( writeUtf8Bom( file ) );
	CStringA line;
	for( size_t i = 0; i < length; i++ )
	{
		const sSegment& seg = segments[ i ];

		if( timestamps )
		{
			line = "[";
			printTime( line, seg.time.begin );
			line += " --> ";
			printTime( line, seg.time.end );
			line += "]  ";
		}
		else
			line = "";

		line += skipBlank( seg.text );
		line += "\r\n";
		CHECK( write( file, line ) );
	}
	return S_OK;
}

HRESULT TranscribeDlg::writeSubRip( const sSegment* const segments, const size_t length, const size_t dupLines , CAtlFile& file )
{
	if (segments == nullptr) {
		logWarning(u8"segments is null, writeSubRip fail!");
		return S_OK;
	}

	CHECK( writeUtf8Bom( file ) );

	logInfo(u8"total lines : %d", length);
	logInfo(u8"remove dup lines : %d", dupLines);

	if (length - dupLines <= 0) {
		return S_OK;
	}

	CStringA line;
	for( size_t i = 0; i < length - dupLines; i++ )
	{
		const sSegment& seg = segments[ i ];
		char ch = '(';

		if (std::strchr(seg.text, ch)) {
			continue;
		}

		line.Format( "%zu\r\n", i + 1 );
		printTime( line, seg.time.begin, true );
		line += " --> ";
		printTime( line, seg.time.end, true );
		line += "\r\n";
		line += skipBlank( seg.text );
		line += "\r\n\r\n";
		CHECK( write( file, line ) );
	}
	return S_OK;
}

HRESULT TranscribeDlg::writeWebVTT( const sSegment* const segments, const size_t length, CAtlFile& file )
{
	CHECK( writeUtf8Bom( file ) );
	CStringA line;
	line = "WEBVTT\r\n\r\n";
	CHECK( write( file, line ) );

	for( size_t i = 0; i < length; i++ )
	{
		const sSegment& seg = segments[ i ];
		line = "";

		printTime( line, seg.time.begin, false );
		line += " --> ";
		printTime( line, seg.time.end, false );
		line += "\r\n";
		line += skipBlank( seg.text );
		line += "\r\n\r\n";
		CHECK( write( file, line ) );
	}
	return S_OK;
}

inline HRESULT TranscribeDlg::newSegmentCallback( Whisper::iContext* ctx, uint32_t n_new )
{
	using namespace Whisper;
	CComPtr<iTranscribeResult> result;
	CHECK( ctx->getResults( transcribeArgs.resultFlags, &result ) );

	// logDebug(reinterpret_cast<const char8_t*>(seg.text));
	sTranscribeLength length;
	CHECK(result->getSize(length));

	const size_t len = length.countSegments;
	size_t i = len - n_new;

	const sSegment* const segments = result->getSegments();
	const sToken* const tokens = result->getTokens();
	CStringA str;
	for (; i < len; i++)
	{
		const sSegment& seg = segments[i];

		for (uint32_t j = 0; j < seg.countTokens; j++)
		{
			const sToken& tok = tokens[seg.firstToken + j];
			if (tok.flags & eTokenFlags::Special)
				continue;
			str += tok.text;
		}
		CStringA oldString = appState.currentResult;
		if (str.Compare(oldString) != 0) {
			appState.duplicatedResults.clear();
			appState.currentResult = str;
			appState.duplicatedSecond = 0;
		}
		else {
			if (appState.duplicatedResults.size() == 0) {
				Whisper::sTimeSpanFields timeBegin = seg.time.begin;
				appState.duplicatedSecond = timeBegin.fullSeconds;
			}
			appState.duplicatedResults.push_back(str);
			logInfo(u8"the same %d: %s", appState.duplicatedResults.size(), cstr(str));
			if (appState.duplicatedResults.size() > 15) {
				logInfo(u8"force stop!");

				Whisper::sTimeSpanFields timeBegin = seg.time.begin;
				CString seconds;
				seconds.Format(_T("%u"), timeBegin.fullSeconds);

				logInfo(u8"the same time final: %u", timeBegin.fullSeconds);

				logInfo(u8"the same time begin: %d", appState.duplicatedSecond);
				
				transcribeStartTime.SetWindowText(seconds);
				transcribeArgs.visualState = eVisualState::Stopping;
			}
		}
	}

	return logNewSegments( result, n_new );
}

HRESULT __cdecl TranscribeDlg::newSegmentCallbackStatic( Whisper::iContext* ctx, uint32_t n_new, void* user_data ) noexcept
{
	TranscribeDlg* dlg = (TranscribeDlg*)user_data;
	return dlg->newSegmentCallback( ctx, n_new );
}

HRESULT __cdecl TranscribeDlg::encoderBeginCallback( Whisper::iContext* ctx, void* user_data ) noexcept
{
	TranscribeDlg* dlg = (TranscribeDlg*)user_data;
	const eVisualState visualState = dlg->transcribeArgs.visualState;
	switch( visualState )
	{
	case eVisualState::Idle:
		return E_NOT_VALID_STATE;
	case eVisualState::Running:
		return S_OK;
	case eVisualState::Stopping:
		return S_FALSE;
	default:
		return E_UNEXPECTED;
	}
}

void TranscribeDlg::onWmClose()
{
	if( GetDlgItem( IDCANCEL ).IsWindowEnabled() )
	{
		EndDialog( IDCANCEL );
		return;
	}

	constexpr UINT flags = MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2;
	const int res = this->MessageBox( L"Transcribe is in progress.\nDo you want to quit anyway?", L"Confirm exit", flags );
	if( res != IDYES )
		return;

	// TODO: instead of ExitProcess(), implement another callback in the DLL API, for proper cancellation of the background task
	ExitProcess( 1 );
}