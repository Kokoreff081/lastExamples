using LightCAD.Model.Entities.Playing;
using LightCAD.Model.Interfaces;
using LightCAD.Model.LT;
using LightCAD.Model.Tools;
using LightControlService.Controllers.Logging;
using LightControlService.Hubs;
using LightControlService.Models.RequestModels;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;

namespace LightControlService.Controllers
{
    public class RendererController : Controller
    {
        private bool startStopAnimation;
        private readonly IScenarioPlayer _scenarioPlayer;
        private readonly ScenarioManager _scenarioManager;
        private ProjectChanger _pchanger;
        private LCHub _lcHub;
        private readonly IScheduleAnalyzer _schedulerAnalyzer;

        public RendererController(IScenarioPlayer scenarioPlayer, ScenarioManager scenarioManager, ProjectChanger pchanger, LCHub lcHub, IScheduleAnalyzer schedulerAnalyzer)
        {
            startStopAnimation = true;
            _scenarioPlayer = scenarioPlayer;
            _scenarioPlayer.FrameChanged += OnFrameChanged;
            _scenarioManager = scenarioManager;
            _pchanger = pchanger;
            _lcHub = lcHub;
            _schedulerAnalyzer = schedulerAnalyzer;
        }

        [HttpPost]
        public void PlayScenario([FromBody] ScenarioNameId snd)
        {
            var scenario = _scenarioManager.GetPrimitives<Scenario>().ToList().First(w => w.Id == snd.ScenarioId && w.Name == snd.ScenarioName);
            if (_scenarioPlayer.InitializeItemId == 0 || _scenarioPlayer.InitializeItemId != scenario.Id)
                _scenarioPlayer.Initialize(scenario);
            if (snd.ElapsedTicks != 0)
                _scenarioPlayer.Rewind((long)snd.ElapsedTicks * 1000);

            _scenarioPlayer.Start();
        }
        private async void OnFrameChanged(object sender, EventArgs e)
        {
            if (startStopAnimation)
            {
                var frame = _scenarioPlayer.Frame;
                var scenario = new WebScenario();
                scenario.lst = new List<WebScenarios>();
                scenario.TotalTicks = _scenarioPlayer.PlayItem.TotalTicks;
                scenario.ElapsedTicks = (float)Math.Round((float)_scenarioPlayer.ElapsedTicks / 1000, 2);
                foreach (var item in frame.GetDictionary())
                {
                    Color frameColor = Color.FromArgb(item.Value[0].Red, item.Value[0].Green, item.Value[0].Blue);
                    string hexColor = string.Format("#{0:X2}{1:X2}{2:X2}", frameColor.R, frameColor.G, frameColor.B);
                    scenario.lst.Add(new WebScenarios() { LampId = item.Key, Color = hexColor });//, Color=item.Value });
                }
                //var sendingFrame = JsonConvert.SerializeObject();
                try
                {
                    await _lcHub.NewFrame(scenario);

                }
                catch (Exception ex)
                {
                    LogManager.GetInstance().ApplicationException(ex.ToString());
                }
            }
        }

        [HttpPost]
        public async void PauseScenario([FromBody] ScenarioNameId snd)
        {
            var frame = _scenarioPlayer.Frame;
            var scenarioToFront = new WebScenario();
            scenarioToFront.TotalTicks = (float)Math.Round((float)_scenarioPlayer.PlayItem.TotalTicks / 1000, 2);
            scenarioToFront.ElapsedTicks = (float)Math.Round((float)_scenarioPlayer.ElapsedTicks / 1000, 2);
            var scenario = _scenarioManager.GetPrimitives<Scenario>().ToList().First(w => w.Id == snd.ScenarioId && w.Name == snd.ScenarioName);
            _scenarioPlayer.Initialize(scenario);

            _scenarioPlayer.Stop();

            scenarioToFront.lst = new List<WebScenarios>();
            foreach (var item in frame.GetDictionary())
            {
                Color frameColor = Color.FromArgb(item.Value[0].Red, item.Value[0].Green, item.Value[0].Blue);
                string hexColor = $"#{frameColor.R:X2}{frameColor.G:X2}{frameColor.B:X2}";
                scenarioToFront.lst.Add(new WebScenarios() { LampId = item.Key, Color = hexColor });//, Color=item.Value });
            }
            try
            {
                await Task.Run(() => _lcHub.NewFrame(scenarioToFront));

            }
            catch (Exception ex)
            {
                LogManager.GetInstance().ApplicationException(ex.ToString());
            }
        }
        [HttpPost]
        public async void StopScenario([FromBody] ScenarioNameId snd)
        {
            var frame = _scenarioPlayer.Frame;
            var scenarioToFront = new WebScenario();
            scenarioToFront.TotalTicks = (float)Math.Round((float)_scenarioPlayer.PlayItem.TotalTicks / 1000, 2);
            scenarioToFront.ElapsedTicks = (float)Math.Round((float)_scenarioPlayer.ElapsedTicks / 1000, 2);
            var scenario = _scenarioManager.GetPrimitives<Scenario>().ToList().First(w => w.Id == snd.ScenarioId && w.Name == snd.ScenarioName);
            _scenarioPlayer.Initialize(scenario);

            _scenarioPlayer.Stop();

            scenarioToFront.lst = new List<WebScenarios>();
            foreach (var item in frame.GetDictionary())
            {
                Color frameColor = Color.Black;
                string hexColor = $"#{frameColor.R:X2}{frameColor.G:X2}{frameColor.B:X2}";
                scenarioToFront.lst.Add(new WebScenarios() { LampId = item.Key, Color = hexColor });//, Color=item.Value });
            }
            try
            {
                await Task.Run(() => _lcHub.NewFrame(scenarioToFront));
                _scenarioPlayer.SwitchOffLamps();

            }
            catch (Exception ex)
            {
                LogManager.GetInstance().ApplicationException(ex.ToString());
            }
        }

        [HttpPost]
        public void StartStopAnimation([FromBody] StartStopScheduler s3)
        {
            _scenarioPlayer.Stop();
            startStopAnimation = s3.action;
            _scenarioPlayer.Start();
        }

        [HttpGet]
        public JsonResult GetCurrentScenario() {
            var resultScenario = new ScenarioNameId();
            resultScenario.IsPlaying = _scenarioPlayer.IsPlay;
            if (_schedulerAnalyzer.CurrentScheduleTask != null)
            {
                resultScenario.ScenarioId = _schedulerAnalyzer.CurrentScheduleTask.ScheduleItem.Scenario.Id;
                resultScenario.ScenarioName = _schedulerAnalyzer.CurrentScheduleTask.ScheduleItem.Scenario.Name;
                resultScenario.TotalTicks = _schedulerAnalyzer.CurrentScheduleTask.ScheduleItem.Scenario.TotalTicks;
                resultScenario.ElapsedTicks = (float)Math.Round((float)_scenarioPlayer.ElapsedTicks / 1000, 2);
            }
            else
            {
                resultScenario = _pchanger.CurrentProject.Scenarios[0];
            }
            return Json(resultScenario);
        }

        [HttpPost]
        public void RewindScenario([FromBody]RewindWebScenario rws)
        {
            try
            {
                if (rws.tick != null)
                {
                    var scenario = _scenarioManager.GetPrimitives<Scenario>().ToList().First(w => w.Id == rws.scenarioId);
                    _scenarioPlayer.Initialize(scenario);
                    _scenarioPlayer.Rewind((long)rws.tick * 1000);
                }
            }
            catch(Exception ex)
            {
                LogManager.GetInstance().ApplicationException(ex.ToString());
            }
        }
    }
}
