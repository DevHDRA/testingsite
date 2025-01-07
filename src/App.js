import React, { useState, useRef, useEffect } from 'react';
import './App.css';
import * as awsChimeSDK from 'amazon-chime-sdk-js';

import { v4 as uuidv4 } from 'uuid';

const axios = require('axios');
const API_URL = process.env.REACT_APP_API_URL;
const ACCOUNT = process.env.REACT_APP_LOAD_TESTING_ACCOUNT;
const SUMADI_TOKEN = process.env.REACT_APP_SUMADI_TOKEN
const createSessionPath = `${ACCOUNT}/${uuidv4()}/`

console.log(API_URL);
console.log(ACCOUNT);

const App = () => {
  const [attendeeId, setAttendeeId] = useState('');
  const [isRecording, setIsRecording] = useState(false)
  const [meetingSessionSt, setMeetingSession] = useState(null);
  const [meetingId, setMeetingId] = useState('')
  const [mediaPipeLine, setMediaPipeLine] = useState('')
  const meetingInformation = useRef({});
  const [proctorParams, setProctorParams] = useState({});

  useEffect(() => {
    const queryParams = new URLSearchParams(window.location.search);
    const params = {};
    queryParams.forEach((value, key) => {
      params[key] = value;
    });
    setProctorParams(params);
  }, []);

  const closeSession = async (event) => {
    event.preventDefault();

    const joinRequest = {
      url: API_URL + 'stop-session',
      method: 'post',
      headers: {
        'Content-Type': 'applications/json',
        'x-source': 'video-recording',
        Authorization: SUMADI_TOKEN,
      },
      data: {
        ...meetingInformation.current,
        meetingCreatedAt: meetingInformation.current.timestamp,
        meetingFinishedAt: new Date(),
        meetingId
      }
    }

    if(meetingSessionSt) {
      const meetingInfo = await axios(joinRequest)
      console.log('Deleted: ', meetingInfo)
      
      meetingSessionSt.audioVideo.stopContentShare();
      await meetingSessionSt.deviceController.destroy();
      meetingSessionSt.audioVideo.stop();
    }
  }

  const handleJoinMeeting = async (event) => {
    event.preventDefault()
    console.log("Joining Meeting", proctorParams);

    const recordingStoragePath = process.env.REACT_APP_STORAGE_PATH;
    meetingInformation.current = {
      'objectPath': createSessionPath,
      proctorId: `${proctorParams.proctorId}`,
      userId: `${proctorParams.userId}`,
      assessmentId: `random-${uuidv4()}`,
      accountId: `random-${uuidv4()}`,
      courseId: `random-${uuidv4()}`,
      timestamp: new Date(),
    };

    console.log(SUMADI_TOKEN);
    const joinRequest = {
      url: API_URL + 'create-session',
      method: 'post',
      headers: {
        'Content-Type': 'applications/json',
        'x-source': 'video-recording',
        Authorization: SUMADI_TOKEN,
      },
      data: meetingInformation.current
    }
    console.log(joinRequest)
    try {
      const meetingInfo = await axios(joinRequest)
      console.log(meetingInfo)
      const configuration = new awsChimeSDK.MeetingSessionConfiguration(
        meetingInfo.data.Meeting,
        meetingInfo.data.Attendee,
      );
      console.log(configuration)
  
      setMeetingId(meetingInfo.data.Meeting.MeetingId);
      setIsRecording(!isRecording);

      console.log('setMeetingId: ', meetingId)
      console.log('setIsRecording: ', isRecording)
      
      const logger = new awsChimeSDK.ConsoleLogger('MyLogger', awsChimeSDK.LogLevel.INFO);
      const deviceController = new awsChimeSDK.DefaultDeviceController(logger);
      const meetingSession = new awsChimeSDK.DefaultMeetingSession(
        configuration,
        logger,
        deviceController,
      );
      setMeetingSession(meetingSession);
      const audioInputDevices = await meetingSession.audioVideo.listAudioInputDevices();
      const audioOutputDevices = await meetingSession.audioVideo.listAudioOutputDevices();
      const videoInputDevices = await meetingSession.audioVideo.listVideoInputDevices();

    // An array of MediaDeviceInfo objects
      audioInputDevices.forEach(mediaDeviceInfo => {
        console.log(`INPUT Device ID: ${mediaDeviceInfo.deviceId} Microphone: ${mediaDeviceInfo.label}`);
      });

      audioOutputDevices.forEach(mediaDeviceInfo => {
        console.log(`OUTPUT Device ID: ${mediaDeviceInfo.deviceId} Output Speaker: ${mediaDeviceInfo.label}`);
      });

      videoInputDevices.forEach(mediaDeviceInfo => {
        console.log(`VIDEO Device ID: ${mediaDeviceInfo.deviceId} video: ${mediaDeviceInfo.label}`);
      });

      await meetingSession.audioVideo.startAudioInput(audioInputDevices[0].deviceId);
      await meetingSession.audioVideo.startVideoInput(videoInputDevices[0].deviceId);

      const audioElement = document.getElementById('audio-element');
      meetingSession.audioVideo.bindAudioElement(audioElement);
      
      const videoElement = document.getElementById('video-element');
      const observer = {
        audioVideoDidStart: () => {
          console.log('Started');
          meetingSession.audioVideo.startLocalVideoTile();
        },
        // videoTileDidUpdate is called whenever a new tile is created or tileState changes.
        videoTileDidUpdate: tileState => {
          // Ignore a tile without attendee ID and other attendee's tile.
        //  console.log('EXECUTING VIDEO TILE UPDATE');
          if (!tileState.boundAttendeeId || !tileState.localTile) {
            return;
          }

          meetingSession.audioVideo.bindVideoElement(tileState.tileId, videoElement);
          const yourAttendeeId = meetingSession.configuration.credentials.attendeeId;

          // tileState.boundAttendeeId is formatted as "attendee-id#content".
          const boundAttendeeId = tileState.boundAttendeeId;
      
          // Get the attendee ID from "attendee-id#content".
          const baseAttendeeId = new awsChimeSDK.DefaultModality(boundAttendeeId).base();
          if (baseAttendeeId === yourAttendeeId) {
            console.log('You called startContentShareFromScreenCapture');
          }
        }
      };

      const contentShareStream = await meetingSession.audioVideo.startContentShareFromScreenCapture();
      awsChimeSDK.DefaultVideoTile.connectVideoStreamToVideoElement(contentShareStream, document.getElementById('screen-video-element'), false);
      meetingSession.audioVideo.addContentShareObserver(observer);
      meetingSession.audioVideo.addObserver(observer);
      meetingSession.audioVideo.start();
    } catch (err) {
      console.log(err)
    }
  }

  return (
    <div>
      <form onSubmit={handleJoinMeeting}>
        <label>
          Attendee ID: <input id="inputName" type="text" value={attendeeId} onChange={(e) => setAttendeeId(e.target.value)} />
        </label>
        <label>
          Meeting Title: <input id="inputMeeting" type="text" />
        </label>
        <input id="authenticate" type="submit" value="Submit" />
      </form>
      <div id="video">

          <video id="video-element" width="350" height="350">

          </video>  

          <video id="screen-video-element" width="400" height="300">

</video>  

          <audio id="audio-element">

          </audio>
       
      </div>
      <button id="button-meeting-leave" onClick={closeSession}>
        Close Session
      </button>
    </div>
  )
}

export default App;
