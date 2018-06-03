using System.Media;

namespace EMecaAddin.Media
{
    public class Sound : ISound
    {
        public void Play()
        {
            const string stopSound = @"C:\Windows\Media\tada.wav";

            if(!System.IO.File.Exists(stopSound))
                return;

            var simpleSound = new SoundPlayer(stopSound);
            simpleSound.Play();
        }
    }
}
