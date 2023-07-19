import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as SpeechSDK from 'microsoft-cognitiveservices-speech-sdk';

export class TextToSpeech2 implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    // state attributes
    private _context : ComponentFramework.Context<IInputs>;
    private _isInitialised : boolean = false;
    private _notifyOutputChanged: () => void;

    // property attributes
    private _text : string = "";
    private _state : string = "waiting";
    private _subscriptionKey : string;
    private _region : string;
    private _language : string;
    private _voice : string = "en-US-ChristopherNeural";
    private _autoSpeak : boolean = false;
    private _playColor : string = "blue";
    private _stopColor : string = "red";
    private _audio : HTMLAudioElement | null = null;

    // ui attributes
    private _container : HTMLDivElement;
    private _buttonDiv : HTMLDivElement;

    /**
     * Empty constructor.
     */
    constructor()
    {
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
        this._context = context;
        this._context.mode.trackContainerResize(true);
        this._container = container;

        // save the notifyOutputChanged
        this._notifyOutputChanged = notifyOutputChanged;
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        //has anything changed?  If not, bug out
        if(this._text === context.parameters.text.raw && 
           this._state === context.parameters.state.raw &&
           this._subscriptionKey === context.parameters.subscriptionKey.raw &&
           this._region === context.parameters.region.raw &&
           this._language === context.parameters.language.raw &&
           this._voice === context.parameters.voice.raw &&
           this._playColor === context.parameters.playColor.raw &&
           this._stopColor === context.parameters.stopColor.raw &&
           this._autoSpeak === context.parameters.autoSpeak.raw) {
            return;
        }

        // update the properties
        this._text = context.parameters.text.raw as string;
        this._state = context.parameters.state.raw as string;
        this._subscriptionKey = context.parameters.subscriptionKey.raw as string;
        this._region = context.parameters.region.raw as string;
        this._language = context.parameters.language.raw as string;
        this._voice = context.parameters.voice.raw as string;
        this._autoSpeak = context.parameters.autoSpeak.raw as boolean;
        this._playColor = context.parameters.playColor.raw as string;
        this._stopColor = context.parameters.stopColor.raw as string;

        // Add code to update control view
        if(!this._isInitialised) {        
            // create the translation div & button
            this._buttonDiv = document.createElement("div");
            this._buttonDiv.id = `button-div`;
            this._buttonDiv.className = `button-div`;
            this._buttonDiv.style.width = `100%`;
            this._buttonDiv.style.height = `100%`;
            this._buttonDiv.style.cursor = `pointer`; 
            this.setPlayButton();
            this._buttonDiv.addEventListener('click', this.buttonPressed.bind(this));
            this._container.appendChild(this._buttonDiv);  
            this._isInitialised = true;
        } else {
            this.setPlayButton();
        }
        if(this._autoSpeak) this.buttonPressed();
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {
            "state" : this._state,
            "autoSpeak" : this._autoSpeak
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }


    public buttonPressed() : void
    {        
        // check that we have text and are not already speaking
        if(this._text === "") {
            return;
        }

        // check if we are already speaking
        if(this._state === "speaking") {
            if(this._audio) {
                this._audio.pause();
                this._audio = null;
                this._state = "idle";
                this._notifyOutputChanged();
                this.setPlayButton();
            }
            return;
        }

        // set state to speaking
        this._state = "speaking";
        this._autoSpeak = false;
        this._notifyOutputChanged();

        // create the async function
        const restAction = async () => {
            const response = await fetch(`https://${this._region}.tts.speech.microsoft.com/cognitiveservices/v1`, {
                method : "post",
                headers: {
                    "Ocp-Apim-Subscription-Key":`${this._subscriptionKey}`,
                    "X-Microsoft-OutputFormat": "riff-24khz-16bit-mono-pcm",
                    "Content-Type": "application/ssml+xml"
                },
                body: `<speak version='1.0' xml:lang='${this._language}'><voice xml:lang='${this._language}' xml:gender='Male' name='${this._voice}'>${this._text}</voice></speak>`
            }
            ).then(result => result.blob()
            ).then(blob => {
                // change to stop button
                this.setStopButton();
                const audioURL = URL.createObjectURL(blob);
                this._audio = new Audio(audioURL);
                this._audio.onended = () => {
                    this._state = "idle"; 
                    this._notifyOutputChanged();
                    this.setPlayButton();
                }
                this._audio.play();
            });
        }

        // do it
        restAction();
    }

    public setPlayButton() {
        this._buttonDiv.innerHTML = `<svg width="${this._context.mode.allocatedWidth}" height="${this._context.mode.allocatedHeight}" viewBox="0 0 1024 1024" fill="none" xmlns="http://www.w3.org/2000/svg"><g clip-path="url(#clip0_626_2)"><circle cx="512" cy="512" r="448" fill="${this._playColor}"/><circle cx="512" cy="512" r="480" stroke="${this._playColor}" stroke-opacity="0.5" stroke-width="64"/><path d="M768 456.574C810.667 481.208 810.667 542.792 768 567.426L432 761.415C389.333 786.049 336 755.257 336 705.99V318.01C336 268.743 389.333 237.951 432 262.585L768 456.574Z" fill="white"/></g><defs><clipPath id="clip0_626_2"><rect width="1024" height="1024" fill="white"/></clipPath></defs></svg>`;
    }

    public setStopButton() {
        this._buttonDiv.innerHTML = `<svg width="${this._context.mode.allocatedWidth}" height="${this._context.mode.allocatedHeight}" viewBox="0 0 1024 1024" fill="none" xmlns="http://www.w3.org/2000/svg"><g clip-path="url(#clip0_236_16)"><circle cx="512" cy="512" r="448" fill="${this._stopColor}"/><circle cx="512" cy="512" r="480" stroke="${this._stopColor}" stroke-opacity="0.5" stroke-width="64"/><rect x="256" y="256" width="512" height="512" rx="64" fill="white"/></g><defs><clipPath id="clip0_236_16"><rect width="1024" height="1024" fill="white"/></clipPath></defs></svg>`;
    }
}