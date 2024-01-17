<template>
    <div class="row sticky-top">
        <h3 class="rowPadd">{{$t("message.projectProjectScenPlayer")}}</h3>
        <div class="col-xl-1 text-center">
            <myselect :options="groups" v-model="selected" @change="onItemSelected" optionLabel="name" optionValue="id" :editable="false">
                <template #value="slotProps">
                    <div class="p-dropdown-code">
                        <span v-if="slotProps.value">{{ slotProps.value.name }}</span>
                        <span v-else>
                            {{ slotProps.placeholder }}
                        </span>
                    </div>
                </template>
            </myselect>
        </div>
        <div class="col-xl-3 text-center">
            <Button class="btn btn-info" @click="startStopAnimation">
                {{$t('message.rendererStartStopAnim')}}
            </Button>
        </div>
        <div class="col-xl-1 text-center">
            <myselect :options="scenarios" v-model="selectedScenario" @change="onScenarioSelect" optionLabel="scenarioName" optionValue="scenarioId" :editable="false">
                <template #value="slotProps">
                    <div class="p-dropdown-code">
                        <span v-if="slotProps.value">{{ slotProps.value.scenarioName }}</span>
                        <span v-else>
                            {{ slotProps.placeholder }}
                        </span>
                    </div>
                </template>
            </myselect>
        </div>
        <div class="col-xl-2 text-center">
            <!--<span class="icon" :class="playActiveSelector()" @click="playScenario(playScen)"></span>
        <span class="icon" :class="pauseActiveSelector()" @click="pauseScenario(playScen)"></span>
        <span class="icon" :class="stopActiveSelector()" @click="stopScenario(playScen)"></span>-->
            <i class="icon" :class="playActiveSelector()" @click="playScenario(playScen)"></i>
            <i class="icon" :class="pauseActiveSelector()" @click="pauseScenario(playScen)"></i>
            <i class="icon" :class="stopActiveSelector()" @click="stopScenario(playScen)"></i>
        </div>
        <div class="col-xl-3 text-center">
            <vue3-slider v-model="sliderVal"
                         :height="10" :handleScale="1.5"
                         :min="sliderMinVal"
                         :max="sliderMaxVal"
                         :alwaysShowHandle="true"
                         :tooltip="tooltipShow"
                         :alwaysShowTooltip="tooltipShow"
                         :tooltipText="'%v'"
                         tooltipColor="#fff" tooltipTextColor="#000"
                         :flipTooltip="true"
                         @drag-end="onSliderValChanged" />
        </div>
        <div class="col-xl-2 text-center">
            <Button class="btn btn-primary" @click="setDefScale">
                <i class="fa fa-refresh"></i> &nbsp;
                {{$t('message.rendererFitToRasters')}}
            </Button>
        </div>
    </div>
    <div class="row">
        <div class="col-xl-12">
            <div class="canvasRenderer">
                <div id="canvas"></div>
            </div>
        </div>
    </div>
    <!--<div class="row">
        <div class="col-xl-12">
            
        </div>
    </div>-->
</template>
<script>
    import { HTTP } from '../../global/commonHttpRequest';
    import { HubConnectionBuilder, LogLevel } from "@microsoft/signalr";
    import slider from "vue3-slider"
    import Button from 'primevue/button';
    import Dropdown from 'primevue/dropdown';

    export default {
        name: "Renderer",
        components: {
            'vue3-slider': slider,
            'myselect': Dropdown,
            'PrimButton': Button,
        },
        props: ['rasters', 'groups', 'scenarioToPlay', 'maxTicks', 'scenarios','playingScenario'],
        data: function () {
            return {
                frameToRender: null,
                connection: null,
                connFlag: false,
                configKonva: {
                    width: 9000,
                    height: 1000,
                },
                selected: { id: 0, name: 'All rasters' },
                defaultSelected: { id: 0, name: 'All rasters' },
                renderRasters: this.rasters,
                stage: null,
                startStopAnim: true,
                scenarioTicks: 0,
                selectedScenario: this.scenarioToPlay,
                playScen: this.playingScenario,
                isPlaying: false,
                playScenarioId: -1,
                sliderVal: 0,
                sliderMinVal: 0,
                sliderMaxVal: 1,
                tooltipShow: true,
                scaleBy: 1.01,
                xScale: 1,
                yScale: 1,
            }
        },
        created: function () {
            this.connection = new HubConnectionBuilder()
                .withUrl("/api/lchub")
                .withAutomaticReconnect()
                .configureLogging(LogLevel.Information)
                .build();

            this.connection.start().then(() => {
                this.connFlag = true;
            }).catch(err => { console.error(err.toString()) });

            HTTP.get('/Renderer/GetCurrentScenario')
                .then(response => {
                    let arr = response.data;
                    console.log(arr);
                    this.sliderVal = arr.elapsedTicks;
                    this.playScen = arr;
                    this.sliderMaxVal = arr.totalTicks;
                })
                .catch(error => { console.log(error); })
        },
        watch: {
            renderRasters: {
                handler() {
                    console.log(this.rasters);
                    this.renderRasters = this.rasters
                },
		        immediate: true 
            },
            maxTicks: {
                handler() {
                    console.log(this.scenarioToPlay);
                    //this.maxTicks = this.scenarioToPlay.totalTicks;
                },
                immediate: true 
            },
        },
        mounted() {
            setTimeout(() => {
                console.log(this.scenarioToPlay);
                this.renderRasters = this.rasters;
                this.drawRastersInitial(this.renderRasters);
                this.selectedScenario = this.scenarioToPlay;
            }, 2000);

            this.connection.on('NewFrame', (frame) => {
                console.log(frame);
                let that = this;
                this.scenarioTicks = frame.elspsedTicks;
                that.sliderVal = frame.elapsedTicks;
                this.frameToRender = frame;
                let stageRef = that.stage.getStage();
                let layer = stageRef.children[0];
                let groups = layer.children;
                for (let i = 0; i < groups.length; i++) {
                    let groupItem = groups[i];
                    for (let j = 0; j < groupItem.children.length; j++) {
                        let shape = groupItem.children[j];
                        let item = this.frameToRender.lst.find(elem => 'Lamp_' + elem.lampId == shape.attrs.id);
                        shape.fill(item.color);
                    }
                }
            });
            this.connection.on('ProjectChanged', (newProject) => {
                let that = this;
                this.stage.destroyChildren();
                that.xScale = 1;
                that.yScale = 1;
                this.redrawRasters(newProject.rasters);
                
            });

            this.connection.on('TasksChanged', (info) => {
                let that = this;
                this.sliderVal = 0;
            });

        },
        methods: {
            drawRastersInitial: function (rasters) {
                console.log(rasters);
                this.stage = new Konva.Stage({
                    container: 'canvas',
                    width: this.configKonva.width,
                    height: this.configKonva.height,
                    draggable: true,
                });
                this.stage.on('wheel', (e) => {
                    // stop default scrolling
                    e.evt.preventDefault();
                    console.log(e.evt.deltaY);
                    var oldScale = this.stage.scaleX();
                    var pointer = this.stage.getPointerPosition();

                    var mousePointTo = {
                        x: (pointer.x - this.stage.x()) / oldScale,
                        y: (pointer.y - this.stage.y()) / oldScale,
                    };

                    // how to scale? Zoom in? Or zoom out?
                    let direction = e.evt.deltaY > 0 ? -1 : 1;

                    // when we zoom on trackpad, e.evt.ctrlKey is true
                    // in that case lets revert direction
                    if (e.evt.ctrlKey) {
                        direction = -direction;
                    }

                    var newScale = direction > 0 ? oldScale * this.scaleBy : oldScale / this.scaleBy;

                    this.stage.scale({ x: newScale, y: newScale });

                    var newPos = {
                        x: pointer.x - mousePointTo.x * newScale,
                        y: pointer.y - mousePointTo.y * newScale,
                    };
                    this.stage.position(newPos);
                });
                var Layer = new Konva.Layer();
                let totalWidth = 0;
                let totalHeight = 0;
                let xPosGroup = rasters[0].id;
                for (let i = 0; i < rasters.length; i++) {
                    let raster = rasters[i];
                    if (raster.dimensionX * 30 > totalWidth)
                        totalWidth += raster.dimensionX * 30;
                    if (raster.dimensionY * 30 > totalHeight)
                        totalHeight += raster.dimensionY * 30;

                    let group = new Konva.Group({
                        y: raster.id * 30,
                        x: xPosGroup,
                        width: raster.dimensionX * 30,
                        heigth: raster.dimensionY * 30,
                        name: raster.name,
                        id: 'Group_' + raster.id,
                        visible: true,
                        stroke: 'yellow',
                        strokeWidth: 4

                    });
                    for(let j = 0; j < raster.projections.length; j++) {
                        let projection = raster.projections[j];
                        let xPosition = projection.rasterX * (projection.width * 3);
                        let box = new Konva.Rect({
                            x: xPosition,
                            y: projection.rasterY * projection.height * 3,
                            id: 'Lamp_' + projection.lampId,
                            fill: projection.color,
                            width: projection.width * 3,
                            height: projection.height * 3,
                            stroke: 'white',
                            strokeWidth: 2,
                        });
                        group.add(box);
                    }
                    Layer.add(group);
                }
                this.stage.add(Layer);

                if (totalWidth > this.configKonva.width || totalHeight > this.configKonva.height) {
                    if (totalWidth > this.configKonva.width && totalHeight > this.configKonva.height)
                    {
                        this.xScale = this.configKonva.width / totalWidth;
                        this.yScale = this.configKonva.height / totalHeight;
                        this.scaleBy = (this.xScale * this.yScale) * 1000;
                    }
                    else if (totalWidth > this.configKonva.width) {
                        this.xScale = this.configKonva.width / totalWidth;
                        this.yScale = this.configKonva.width / totalWidth;
                        this.scaleBy = (totalWidth / 10) / this.configKonva.width;
                    }
                    else {
                        this.yScale = this.configKonva.height / totalHeight;
                        this.xScale = this.configKonva.height / totalHeight;
                        this.scaleBy = (totalHeight / 10) / this.configKonva.height;
                    }
                    let stageScale = { x: this.xScale, y: this.yScale };
                    this.stage.scale(stageScale);
                }
                else {
                    let stageScale = { x: 1, y: 1 };
                    this.stage.scale(stageScale);
                    this.scaleBy = 1.01;
                }
            },
            redrawRasters: function (rasters) {
                this.stage.clear();
                let layer = new Konva.Layer();
                let totalWidth = 0;
                let totalHeight = 0;
                let xPosGroup = rasters[0].id;
                for (let i = 0; i < rasters.length; i++) {
                    let raster = rasters[i];
                    if (raster.dimensionX * 30 > totalWidth)
                        totalWidth += raster.dimensionX * 30;
                    if (raster.dimensionY * 30 > totalHeight)
                        totalHeight += raster.dimensionY * 30;

                    var group = new Konva.Group({
                        y: raster.id * 30,
                        x: xPosGroup,
                        width: raster.dimensionX * 30,
                        heigth: raster.dimensionY * 30,
                        offset: { offsetX: raster.id * 10, offsetY: raster.id * 10 },
                        name: raster.name,
                        id: 'Group_' + raster.id,
                        visible: true
                    });
                    for (let j = 0; j < raster.projections.length; j++) {
                        let projection = raster.projections[j];
                        let xPosition = projection.rasterX * (projection.width * 3);
                        
                        var box = new Konva.Rect({
                            x: xPosition,
                            y: projection.rasterY * projection.height * 3,
                            id: 'Lamp_' + projection.lampId,
                            fill: projection.color,
                            width: projection.width * 3,
                            height: projection.height * 3,
                            stroke: 'white',
                            strokeWidth: 2,
                        });
                        group.add(box);
                    }
                    layer.add(group);
                }
                this.stage.add(layer);
                if (totalWidth > this.configKonva.width || totalHeight > this.configKonva.height) {
                    if (totalWidth > this.configKonva.width && totalHeight > this.configKonva.height) {
                        this.xScale = this.configKonva.width / totalWidth;
                        this.yScale = this.configKonva.height / totalHeight;
                        this.scaleBy = (this.xScale * this.yScale) * 1000;
                    }
                    else if (totalWidth > this.configKonva.width) {
                        this.xScale = this.configKonva.width / totalWidth;
                        this.yScale = this.configKonva.width / totalWidth;
                        this.scaleBy = (totalWidth / 10) / this.configKonva.width;
                    }
                    else {
                        this.yScale = this.configKonva.height / totalHeight;
                        this.xScale = this.configKonva.height / totalHeight;
                        this.scaleBy = (totalHeight / 10) / this.configKonva.height;
                    }
                    let stageScale = { x: this.xScale, y: this.yScale };
                    this.stage.scale(stageScale);
                }
                else {
                    let stageScale = { x: 1, y: 1 };
                    this.stage.scale(stageScale);
                    this.scaleBy = 1.01;
                }
            },
            startStopAnimation: function () {
                let headers = { 'Content-Type': 'application/json' };
                console.log(this.startStopAnim);
                this.startStopAnim = !this.startStopAnim;
                console.log(this.startStopAnim);
              
                let StartStopScheduler = { action: this.startStopAnim };
                HTTP.post('/Renderer/StartStopAnimation', StartStopScheduler, { headers });
            },
            playScenario: function (scenario) {
                if (!scenario.isPlaying) {
                    this.isPlaying = true;
                    this.playScenarioId = scenario.scenarioId;
                    scenario.isPlaying = true;
                    let ScenarioNameId = { "scenarioId": scenario.scenarioId, "scenarioName": scenario.scenarioName, "elapsedTicks": this.sliderVal };
                    let headers = { 'Content-Type': 'application/json' };
                    HTTP.post('/Renderer/PlayScenario', ScenarioNameId, { headers });
                }
            },
            pauseScenario: function (scenario) {
                if (scenario.isPlaying) {
                    console.log(scenario.scenarioId, scenario.scenarioName);
                    scenario.isPlaying = false;
                    this.isPlaying = false;
                    let ScenarioNameId = { "scenarioId": scenario.scenarioId, "scenarioName": scenario.scenarioName };
                    let headers = { 'Content-Type': 'application/json' };
                    HTTP.post('/Renderer/PauseScenario', ScenarioNameId, { headers });
                }
            },
            stopScenario: function (scenario) {
                if (scenario.isPlaying) {
                    console.log(scenario.scenarioId, scenario.scenarioName);
                    scenario.isPlaying = false;
                    this.isPlaying = false;
                    let ScenarioNameId = { "scenarioId": scenario.scenarioId, "scenarioName": scenario.scenarioName };
                    let headers = { 'Content-Type': 'application/json' };
                    HTTP.post('/Renderer/StopScenario', ScenarioNameId, { headers });
                    this.playScenarioId = -1;
                    this.sliderVal = 0;
                }
            },
            playActiveSelector: function () {
                if (this.isPlaying)
                    return 'icon-playScenario-inactive';
                else
                    return 'icon-playScenario';
            },
            pauseActiveSelector: function () {
                if (this.playScenarioId == -1)
                    return 'icon-pauseScenario-inactive';
                else {
                    if (this.isPlaying) {
                        return 'icon-pauseScenario';
                    }
                    else
                        return 'icon-pauseScenario-inactive';
                }
            },
            stopActiveSelector: function () {
                if (this.playScenarioId == -1)
                    return 'icon-stopScenario-inactive';
                else {
                    if (this.isPlaying) {
                        return 'icon-stopScenario';
                    }
                    else
                        return 'icon-stopScenario-inactive';
                }
            },
            onSliderValChanged: function () {
                console.log(this.sliderVal);
                let rws = { tick: this.sliderVal, scenarioId: this.playScen.scenarioId };
                let headers = { 'Content-Type': 'application/json' };
                HTTP.post('/Renderer/RewindScenario', rws, { headers });
            },
            onItemSelected: function () {
                console.log(this.selected, this.groups);
                let findedRaster = this.groups.find(elem => elem.id == this.selected);
                console.log(findedRaster);
                this.selected = findedRaster;
                console.log(this.selected);
                let stageRef = this.stage.getStage();
                let layer = stageRef.children[0];
                let groupObjects = layer.children;
                if (this.selected.name !== 'All rasters') {
                    for (let i = 0; i < groupObjects.length; i++) {

                        let group = groupObjects[i];
                        if (group.attrs.name != this.selected.name)
                            group.visible(false);
                        else
                            group.visible(true);
                    }
                }
                else {
                    for (let i = 0; i < groupObjects.length; i++) {
                        let group = groupObjects[i];
                        group.visible(true);
                    }
                }
            },
            onScenarioSelect: function () {
                let findedScen = this.scenarios.find(f => f.scenarioId == this.selectedScenario);
                this.selectedScenario = findedScen;
                console.log(findedScen);
                this.playScen = {
                    scenarioId: findedScen.scenarioId,
                    scenarioName: findedScen.scenarioName,
                    isPlaying: findedScen.isPlaying,
                    scenarioTime: findedScen.scenarioTime,
                    totalTicks: findedScen.totalTicks,
                };
                console.log(this.playScen);
                this.sliderMaxVal = this.playScen.totalTicks;
            },
            setDefScale: function () {
                /*let stageScale = { x: this.xScale, y: this.yScale };
                console.log(stageScale);
                this.stage.scale(stageScale);*/
                // do we need padding?
                let padding = 10;
                let layer = this.stage.children[0];
                // get bounding rectangle
                let box = layer.getClientRect({ relativeTo: this.stage });

                this.stage.setAttrs({
                    x: box.x * this.xScale,
                    y: box.y * this.yScale,
                    scaleX: this.xScale,
                    scaleY: this.yScale
                });
            }
        },
    }

</script>